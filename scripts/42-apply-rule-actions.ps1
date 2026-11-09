[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$Database,
    [Parameter(Mandatory)] [string]$RuleActions,
    [Parameter(Mandatory)] [string]$ExpectedRuleActionsSha256,
    [Parameter(Mandatory)] [string]$ObjectManifest,
    [Parameter(Mandatory)] [string]$ExpectedManifestSha256,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-rule-actions')),
    [switch]$Apply,
    [string]$ConfirmationToken
)
<#
    Executes the column-level boundary rules exported by script 41
    (rule-actions.csv) against the frozen delete manifest, before script 11.

      CLEAR_FIELD  the reference is set to an empty value.  A composite field
                   keeps its _TYPE and _RTRef, so the value becomes an empty
                   reference of the same type (Организации.ПустаяСсылка), not
                   Неопределено.
      CLEAR_ROW    the tabular or register row is deleted; the owner stays.

    Only references to objects of the manifest are touched, and rows whose
    owner is itself in the manifest are skipped (script 11 deletes them).
    Without -Apply the script only counts.  -Apply runs every action in one
    transaction and requires the token CLEAR-RULE-ACTIONS.
#>
$ErrorActionPreference = 'Stop'
$zeroId = '0x' + ('0' * 32)

function Quote-SqlIdentifier([string]$Name) {
    if ($Name -notmatch '^_[A-Za-z0-9_]+$') { throw "Unsafe SQL identifier: $Name" }
    return '[' + $Name + ']'
}
function Get-TypeCode([string]$TypeTable) {
    if ($TypeTable -notmatch '^[A-Za-z]+(?<code>[0-9]+)$') { throw "Invalid type token: $TypeTable" }
    return [int]$Matches.code
}
function Assert-Hash([string]$Path, [string]$Expected, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label was not found: $Path" }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $Expected.ToLowerInvariant()) { throw "$Label SHA-256 changed: $actual" }
    return $actual
}
function Invoke-Scalar([System.Data.SqlClient.SqlConnection]$Connection, [System.Data.SqlClient.SqlTransaction]$Transaction, [string]$Sql) {
    $command = $Connection.CreateCommand()
    $command.Transaction = $Transaction
    $command.CommandTimeout = 3600
    $command.CommandText = "SET NOCOUNT ON;`n$Sql"
    try { return $command.ExecuteScalar() } finally { $command.Dispose() }
}

if ($Apply -and $ConfirmationToken -ne 'CLEAR-RULE-ACTIONS') { throw '-Apply requires -ConfirmationToken CLEAR-RULE-ACTIONS.' }
$actionsHash = Assert-Hash $RuleActions $ExpectedRuleActionsSha256 'Rule actions'
$manifestHash = Assert-Hash $ObjectManifest $ExpectedManifestSha256 'Object manifest'
$actions = @(Import-Csv -LiteralPath $RuleActions)
if ($actions.Count -eq 0) { throw 'Rule actions file is empty.' }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$connection = [System.Data.SqlClient.SqlConnection]::new(
    "Server=$Server;Database=$Database;Integrated Security=True;Encrypt=False;Application Name=Codex-RuleActions")
$started = Get-Date
try {
    $connection.Open()
    if ([string](Invoke-Scalar $connection $null 'SELECT DB_NAME();') -ne $Database) { throw "Connected to an unexpected database. Expected $Database." }

    [void](Invoke-Scalar $connection $null 'CREATE TABLE #TargetObjects(type_code int NOT NULL, id binary(16) NOT NULL, PRIMARY KEY(type_code,id)); SELECT 1;')
    $table = [System.Data.DataTable]::new()
    [void]$table.Columns.Add('type_code', [int])
    [void]$table.Columns.Add('id', [byte[]])
    foreach ($row in Import-Csv -LiteralPath $ObjectManifest) {
        $bytes = [byte[]]::new(16)
        for ($i = 0; $i -lt 16; $i++) { $bytes[$i] = [Convert]::ToByte(([string]$row.id_hex).Substring(2 + 2*$i, 2), 16) }
        [void]$table.Rows.Add([int]$row.type_code, $bytes)
    }
    $bulk = [System.Data.SqlClient.SqlBulkCopy]::new($connection)
    try { $bulk.DestinationTableName = '#TargetObjects'; $bulk.BulkCopyTimeout = 900; $bulk.WriteToServer($table) } finally { $bulk.Dispose() }

    $plans = foreach ($action in $actions) {
        $target = Quote-SqlIdentifier ([string]$action.physical_table)
        $column = Quote-SqlIdentifier ([string]$action.column)
        # SQL Server 2014: no STRING_AGG.
        $columns = @(([string](Invoke-Scalar $connection $null "SELECT STUFF((SELECT N','+name FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.$([string]$action.physical_table)') FOR XML PATH(N'')),1,1,N'');")) -split ',')
        if ($columns -notcontains [string]$action.column) { throw "Column $($action.physical_table).$($action.column) does not exist." }
        if ($action.type_column) {
            $typeColumn = Quote-SqlIdentifier ([string]$action.type_column)
            $reference = "t.type_code=CONVERT(int,CONVERT(varbinary(4),src.$typeColumn)) AND t.id=src.$column"
            if ($action.target_type_code) { $reference += " AND src.$typeColumn=CONVERT(binary(4),$([int]$action.target_type_code))" }
        }
        else {
            $fixedCode = Get-TypeCode ([string]$action.target_table)
            if ($action.target_type_code -and [int]$action.target_type_code -ne $fixedCode) { continue }
            $reference = "t.type_code=$fixedCode AND t.id=src.$column"
        }
        $predicate = "EXISTS(SELECT 1 FROM #TargetObjects t WHERE $reference)"
        # Skip rows that disappear with their owner anyway.
        $owner = [string]$action.owner_column
        $ownerType = ([string]$action.physical_table).TrimStart('_') -replace '_VT[0-9]+','' -replace 'X[0-9]+$',''
        if ($owner -and $columns -contains $owner -and $ownerType -match '^[A-Za-z]+[0-9]+$') {
            $predicate += " AND NOT EXISTS(SELECT 1 FROM #TargetObjects o WHERE o.type_code=$(Get-TypeCode $ownerType) AND o.id=src.$(Quote-SqlIdentifier $owner))"
        }
        $sql = switch ([string]$action.action) {
            'CLEAR_FIELD' { "UPDATE src SET $column=$zeroId FROM dbo.$target src WHERE $predicate;" }
            'CLEAR_ROW' { "DELETE src FROM dbo.$target src WHERE $predicate;" }
            default { throw "Unsupported action $($action.action)." }
        }
        # A tabular part must keep its line numbers 1..N without gaps, or 1C
        # "Тестирование и исправление" reports "Нарушена нумерация строк".
        $renumber = $null
        $lineNo = @($columns | Where-Object { $_ -match '^_LineNo[0-9]+$' }) | Select-Object -First 1
        if ($action.action -eq 'CLEAR_ROW' -and $owner -and $columns -contains $owner -and $lineNo) {
            $qOwner = Quote-SqlIdentifier $owner
            $qLine = Quote-SqlIdentifier $lineNo
            $sql = "CREATE TABLE #Renumber(owner_id binary(16) NOT NULL PRIMARY KEY); " +
                   "INSERT #Renumber(owner_id) SELECT DISTINCT src.$qOwner FROM dbo.$target src WHERE $predicate; " + $sql
            $renumber = "WITH n AS (SELECT src.$qLine AS line_no, ROW_NUMBER() OVER (PARTITION BY src.$qOwner ORDER BY src.$qLine) AS new_no " +
                        "FROM dbo.$target src JOIN #Renumber r ON r.owner_id=src.$qOwner) " +
                        "UPDATE n SET line_no=new_no WHERE line_no<>new_no; DROP TABLE #Renumber;"
        }
        [pscustomobject]@{ rule=$action.rule; action=$action.action; table=$action.physical_table; column=$action.column
                           count_sql="SELECT COUNT_BIG(*) FROM dbo.$target src WHERE $predicate;"; dml=$sql; renumber=$renumber; rows=[int64]0 }
    }

    foreach ($plan in $plans) { $plan.rows = [int64](Invoke-Scalar $connection $null $plan.count_sql) }
    $applied = [int64]0
    if ($Apply) {
        $transaction = $connection.BeginTransaction()
        try {
            foreach ($plan in $plans) {
                $command = $connection.CreateCommand()
                $command.Transaction = $transaction
                $command.CommandTimeout = 3600
                $command.CommandText = "SET XACT_ABORT ON; SET NOCOUNT ON; $($plan.dml) SELECT CONVERT(bigint,@@ROWCOUNT);"
                $done = [int64]$command.ExecuteScalar()
                $command.Dispose()
                if ($done -ne $plan.rows) { throw "$($plan.table).$($plan.column): planned $($plan.rows), changed $done." }
                $applied += $done
                if ($plan.renumber) {
                    $command = $connection.CreateCommand()
                    $command.Transaction = $transaction
                    $command.CommandTimeout = 3600
                    $command.CommandText = "SET XACT_ABORT ON; SET NOCOUNT ON; $($plan.renumber)"
                    [void]$command.ExecuteNonQuery()
                    $command.Dispose()
                }
            }
            $transaction.Commit()
        }
        catch { try { $transaction.Rollback() } catch {}; throw }
        $left = [int64](($plans | ForEach-Object { [int64](Invoke-Scalar $connection $null $_.count_sql) } | Measure-Object -Sum).Sum)
        if ($left -ne 0) { throw "Postcheck: $left references to the manifest remain after Apply." }
    }

    $plans | Select-Object rule,action,table,column,rows | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'rule-actions-plan.csv') -NoTypeInformation -Encoding UTF8
    $summary = [ordered]@{
        generated_at_local=(Get-Date).ToString('s'); server=$Server; database=$Database
        read_only=(-not $Apply); dml_executed=[bool]$Apply
        rule_actions=(Resolve-Path -LiteralPath $RuleActions).Path; rule_actions_sha256=$actionsHash
        object_manifest=(Resolve-Path -LiteralPath $ObjectManifest).Path; object_manifest_sha256=$manifestHash
        planned_rows=[int64](($plans | Measure-Object -Property rows -Sum).Sum); changed_rows=$applied
        by_action=@($plans | Group-Object action | ForEach-Object { [ordered]@{ action=$_.Name; rows=[int64](($_.Group | Measure-Object rows -Sum).Sum) } })
        duration_seconds=[Math]::Round(((Get-Date)-$started).TotalSeconds,1)
    }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
    $summary | ConvertTo-Json -Depth 5
}
finally { $connection.Dispose() }
