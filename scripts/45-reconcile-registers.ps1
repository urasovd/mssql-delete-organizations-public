[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$SourceDatabase,
    [Parameter(Mandatory)] [string]$TargetDatabase,
    # Kept organizations by reference (see lib/Organizations.ps1), one or more,
    # or by code where the catalog has one.
    [Alias('OrganizationId')] [string[]]$Organization,
    [string[]]$OrganizationCode,
    [Parameter(Mandatory)] [string]$CatalogDirectory,
    [Parameter(Mandatory)] [datetime]$PeriodTo,
    [datetime]$PeriodFrom,
    [string[]]$RegisterNames = @('*'),
    [string]$OrganizationFieldName = 'Организация',
    [string]$ExceptionsPath = (Join-Path $PSScriptRoot '..\policies\reconcile-exceptions.csv'),
    [ValidateRange(0,3000)] [int]$DateOffsetYears = 2000,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-reconcile'))
)
<#
    Read-only reconciliation of register movements of the kept organizations between
    a source copy (before deletion) and a target copy (after deletion), up to
    a closed period.  Nothing is taken from totals tables, so the result does
    not depend on whether totals are used or recalculated.

    For every register whose logical name matches -RegisterNames (accounting,
    accumulation, information and calculation registers; "*" is a name pattern) and for
    every physical movement table (base and X extensions) the script compares
    the rows of the organization with period <= -PeriodTo:

      rows       multiset of BINARY_CHECKSUM over all columns: rows missing in
                 the target and extra rows in the target.  When they differ and
                 the table has varbinary(max) columns, the rows are compared
                 again by content: a .dt load replaces some value-storage
                 values with a 64-byte STORHDR header pointing to BinaryData,
                 so the same value has different bytes in the column;
      resources  SUM of every numeric resource column (amounts, quantities,
                 currency amounts);
      accounts   for accounting registers, per account: balance at -PeriodTo
                 and debit/credit turnover for -PeriodFrom..-PeriodTo of the
                 field named "Сумма" (an OSV computed from movements);
      extra dims accounting extra-dimension rows (AccRgED) of those movements;
      act. period  calculation-register actual action period rows (CRgActP) of
                 those movements: the table has no organization field.

    Dates in 1C databases on MS SQL are stored with an offset of 2000 years.
    Result PASS only when every compared table has 0 missing, 0 extra rows and
    equal sums.  A register listed in -ExceptionsPath (a known rule effect or
    data that 1C rewrites at runtime) is reported as EXPECTED with its reason
    and does not fail the run.
#>
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\Organizations.ps1')
$organizationFilter = Get-OrganizationFilter $Organization $OrganizationCode
if (-not $PSBoundParameters.ContainsKey('PeriodFrom')) { $PeriodFrom = Get-Date -Year $PeriodTo.Year -Month $PeriodTo.Month -Day 1 }

function Quote([string]$Name) {
    if ($Name -notmatch '^[A-Za-z0-9_]+$') { throw "Unsafe SQL identifier: $Name" }
    return '[' + $Name + ']'
}
function Invoke-Table([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql) {
    $command = $Connection.CreateCommand()
    $command.CommandTimeout = 3600
    $command.CommandText = "SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`n$Sql"
    try {
        $adapter = [System.Data.SqlClient.SqlDataAdapter]::new($command)
        $table = [System.Data.DataTable]::new()
        [void]$adapter.Fill($table)
        $adapter.Dispose()
        Write-Output -NoEnumerate $table
    }
    finally { $command.Dispose() }
}
function Get-DbDate([datetime]$Date) { return $Date.AddYears($DateOffsetYears).ToString('yyyy-MM-ddTHH:mm:ss') }
function Invoke-Reader([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql) {
    $command = $Connection.CreateCommand()
    $command.CommandTimeout = 3600
    $command.CommandText = "SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`n$Sql"
    # A reader is enumerable: without -NoEnumerate the pipeline would consume it.
    Write-Output -NoEnumerate $command.ExecuteReader()
}
# Multiset of the rows of a table with varbinary(max) values replaced by the
# SHA-256 of their content; a STORHDR value is hashed from its BinaryData chunks.
function Get-ContentRowCounts([System.Data.SqlClient.SqlConnection]$Connection, [string]$Db, [string]$Table, [string]$Filter, [string[]]$PlainColumns, [string[]]$LobColumns) {
    $header = '0x53544F52484452'
    $external = [System.Collections.Generic.Dictionary[string,string]]::new()
    foreach ($column in $LobColumns) {
        $lob = "s.$(Quote $column)"
        $reader = Invoke-Reader $Connection @"
IF OBJECT_ID(N'$Db.dbo.BinaryData') IS NOT NULL
SELECT d.f_key, d.f_data FROM $Db.dbo.BinaryData d
WHERE d.f_key IN (SELECT CONVERT(binary(16),SUBSTRING($lob,17,16)) FROM $Db.dbo.$(Quote $Table) s WHERE $Filter AND DATALENGTH($lob)=64 AND SUBSTRING($lob,1,7)=$header)
ORDER BY d.f_key, d.f_off, d.f_num;
"@
        try {
            $current = $null
            $hash = $null
            while ($reader.Read()) {
                $key = [Convert]::ToHexString([byte[]]$reader[0])
                if ($key -ne $current) {
                    if ($hash) { $external[$current] = [Convert]::ToHexString($hash.GetHashAndReset()) }
                    $hash = [System.Security.Cryptography.IncrementalHash]::CreateHash([System.Security.Cryptography.HashAlgorithmName]::SHA256)
                    $current = $key
                }
                $hash.AppendData([byte[]]$reader[1])
            }
            if ($hash) { $external[$current] = [Convert]::ToHexString($hash.GetHashAndReset()) }
        }
        finally { $reader.Dispose() }
    }

    $counts = [System.Collections.Generic.Dictionary[string,int64]]::new()
    $externalValues = [int64]0
    $plain = ($PlainColumns | ForEach-Object { "s.$(Quote $_)" }) -join ','
    $lobs = ($LobColumns | ForEach-Object { "s.$(Quote $_)" }) -join ','
    $reader = Invoke-Reader $Connection "SELECT BINARY_CHECKSUM($plain) AS h, $lobs FROM $Db.dbo.$(Quote $Table) s WHERE $Filter;"
    try {
        while ($reader.Read()) {
            $parts = [System.Collections.Generic.List[string]]::new()
            $parts.Add([string]$reader[0])
            for ($i = 1; $i -lt $reader.FieldCount; $i++) {
                if ($reader.IsDBNull($i)) { $parts.Add('null'); continue }
                [byte[]]$value = $reader[$i]
                if ($value.Length -eq 64 -and [Text.Encoding]::ASCII.GetString($value, 0, 7) -eq 'STORHDR') {
                    $externalValues++
                    $key = [Convert]::ToHexString($value, 16, 16)
                    # A key without chunks never matches, so the row stays a difference.
                    $parts.Add($(if ($external.ContainsKey($key)) { $external[$key] } else { "no-binary-data:$Db" + ":$key" }))
                }
                else { $parts.Add([Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($value))) }
            }
            $row = $parts -join '|'
            $count = [int64]0
            [void]$counts.TryGetValue($row, [ref]$count)
            $counts[$row] = $count + 1
        }
    }
    finally { $reader.Dispose() }
    return [pscustomobject]@{ counts=$counts; external_values=$externalValues }
}

$src = Quote $SourceDatabase
$tgt = Quote $TargetDatabase
$typeNames = @{}
foreach ($row in Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'type-names.csv')) { $typeNames[$row.type_table] = $row.logical_name }
$fieldNames = @{}
foreach ($row in Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'field-names.csv')) { $fieldNames[$row.field] = $row.logical_name }
$catalog = @(Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'reference-catalog.csv'))
$orgType = @($typeNames.GetEnumerator() | Where-Object { $_.Value -eq 'Организации' -and $_.Name -match '^Reference[0-9]+$' })[0].Name
$exceptions = @{}
if ($ExceptionsPath -and (Test-Path -LiteralPath $ExceptionsPath -PathType Leaf)) {
    foreach ($row in Import-Csv -LiteralPath $ExceptionsPath -Delimiter ';') { $exceptions[[string]$row.register] = "$($row.kind): $($row.reason)" }
}

$registers = @($typeNames.GetEnumerator() | Where-Object {
    $_.Name -match '^(AccRg|AccumRg|InfoRg|CRg)[0-9]+$' -and $(foreach ($p in $RegisterNames) { if ($_.Value -like $p) { $true; break } })
} | Sort-Object Value)
if ($registers.Count -eq 0) { throw "No register matches $($RegisterNames -join ', ')." }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$connection = [System.Data.SqlClient.SqlConnection]::new("Server=$Server;Database=$TargetDatabase;Integrated Security=True;Encrypt=False;Application Name=Codex-Reconcile")
$started = Get-Date
try {
    $connection.Open()
    $orgWhere = "$(Get-OrganizationFilter $Organization $OrganizationCode 's' | ForEach-Object sql) OR $(Get-OrganizationFilter $Organization $OrganizationCode 't' | ForEach-Object sql)"
    $orgRows = Invoke-Table $connection "SELECT sys.fn_varbintohexstr(s._IDRRef) AS src_id, sys.fn_varbintohexstr(t._IDRRef) AS tgt_id FROM $src.dbo.$(Quote ('_' + $orgType)) s FULL JOIN $tgt.dbo.$(Quote ('_' + $orgType)) t ON t._IDRRef=s._IDRRef WHERE $orgWhere;"
    $mismatched = @($orgRows.Rows | Where-Object { [string]$_.src_id -ne [string]$_.tgt_id })
    if ($orgRows.Rows.Count -ne $organizationFilter.expected -or $mismatched.Count) { throw "Organizations '$($organizationFilter.label)' must each exist once with the same ID in both databases." }
    $orgIds = @($orgRows.Rows | ForEach-Object { [string]$_.src_id } | Sort-Object)
    $orgId = $orgIds -join ','
    $dbPeriodTo = Get-DbDate $PeriodTo.Date.AddDays(1).AddSeconds(-1)
    $dbPeriodFrom = Get-DbDate $PeriodFrom.Date

    $tableResults = [System.Collections.Generic.List[object]]::new()
    $resourceResults = [System.Collections.Generic.List[object]]::new()
    $accountResults = [System.Collections.Generic.List[object]]::new()

    foreach ($register in $registers) {
        $type = [string]$register.Name
        $physical = Invoke-Table $connection @"
SELECT t.name FROM $tgt.sys.tables t WHERE t.name=N'_$type' OR t.name LIKE N'[_]$type[X][0-9]%'
INTERSECT
SELECT s.name FROM $src.sys.tables s WHERE s.name=N'_$type' OR s.name LIKE N'[_]$type[X][0-9]%';
"@
        foreach ($tableRow in $physical.Rows) {
            $table = [string]$tableRow.name
            $columns = @((Invoke-Table $connection "SELECT c.name, ty.name AS type_name, c.max_length FROM $tgt.sys.columns c JOIN $tgt.sys.types ty ON ty.user_type_id=c.user_type_id WHERE c.object_id=OBJECT_ID(N'$TargetDatabase.dbo.$table') ORDER BY c.column_id;").Rows)
            $names = @($columns | ForEach-Object { [string]$_.name })
            $orgColumns = @($catalog | Where-Object { $_.physical_table -eq $table -and $_.target_table -eq $orgType -and $_.kind -in @('fixed','composite-single') })
            $orgColumn = @($orgColumns | Where-Object field_name -eq $OrganizationFieldName | Select-Object -First 1).column
            if (-not $orgColumn) { $orgColumn = @($orgColumns | Select-Object -First 1).column }
            if (-not $orgColumn) { continue }
            $filter = "s.$(Quote $orgColumn) IN ($orgId)"
            if ($names -contains '_Period') { $filter += " AND s._Period<='$dbPeriodTo'" }

            $compare = Invoke-Table $connection @"
SELECT BINARY_CHECKSUM(*) AS h INTO #s FROM $src.dbo.$(Quote $table) s WHERE $filter;
SELECT BINARY_CHECKSUM(*) AS h INTO #t FROM $tgt.dbo.$(Quote $table) s WHERE $filter;
WITH a AS (SELECT h, COUNT_BIG(*) AS c FROM #s GROUP BY h), b AS (SELECT h, COUNT_BIG(*) AS c FROM #t GROUP BY h)
SELECT (SELECT COUNT_BIG(*) FROM #s) AS source_rows, (SELECT COUNT_BIG(*) FROM #t) AS target_rows,
       ISNULL(SUM(CASE WHEN ISNULL(a.c,0)>ISNULL(b.c,0) THEN ISNULL(a.c,0)-ISNULL(b.c,0) ELSE 0 END),0) AS missing_in_target,
       ISNULL(SUM(CASE WHEN ISNULL(b.c,0)>ISNULL(a.c,0) THEN ISNULL(b.c,0)-ISNULL(a.c,0) ELSE 0 END),0) AS extra_in_target
FROM a FULL JOIN b ON b.h=a.h;
DROP TABLE #s; DROP TABLE #t;
"@
            $c = $compare.Rows[0]
            $missing = [int64]$c.missing_in_target
            $extra = [int64]$c.extra_in_target
            $externalValues = [int64]0
            $lobColumns = @($columns | Where-Object { $_.type_name -eq 'varbinary' -and [int]$_.max_length -eq -1 } | ForEach-Object { [string]$_.name })
            if (($missing -or $extra) -and $lobColumns.Count) {
                # BINARY_CHECKSUM(*) ignores these types and an explicit list rejects them.
                $plainColumns = @($columns | Where-Object { $_.name -notin $lobColumns -and $_.type_name -notin @('image','text','ntext','xml') } | ForEach-Object { [string]$_.name })
                $sourceContent = Get-ContentRowCounts $connection $src $table $filter $plainColumns $lobColumns
                $targetContent = Get-ContentRowCounts $connection $tgt $table $filter $plainColumns $lobColumns
                $externalValues = $sourceContent.external_values + $targetContent.external_values
                $missing = [int64]0
                $extra = [int64]0
                foreach ($entry in $sourceContent.counts.GetEnumerator()) {
                    $other = [int64]0
                    [void]$targetContent.counts.TryGetValue($entry.Key, [ref]$other)
                    if ($entry.Value -gt $other) { $missing += $entry.Value - $other }
                }
                foreach ($entry in $targetContent.counts.GetEnumerator()) {
                    $other = [int64]0
                    [void]$sourceContent.counts.TryGetValue($entry.Key, [ref]$other)
                    if ($entry.Value -gt $other) { $extra += $entry.Value - $other }
                }
            }

            # Resource sums: numeric columns except service ones.
            $numeric = @($columns | Where-Object { $_.type_name -eq 'numeric' -and $_.name -match '^_Fld[0-9]+(Dt|Ct)?$' -and $_.name -ne '_Fld13373' } | ForEach-Object { [string]$_.name })
            $sumMismatches = 0
            if ($numeric.Count -gt 0) {
                $sumExpr = ($numeric | ForEach-Object { "SUM(CONVERT(decimal(38,4),s.$(Quote $_))) AS $(Quote $_)" }) -join ','
                $srcSums = (Invoke-Table $connection "SELECT $sumExpr FROM $src.dbo.$(Quote $table) s WHERE $filter;").Rows[0]
                $tgtSums = (Invoke-Table $connection "SELECT $sumExpr FROM $tgt.dbo.$(Quote $table) s WHERE $filter;").Rows[0]
                foreach ($column in $numeric) {
                    $a = if ($srcSums.IsNull($column)) { [decimal]0 } else { [decimal]$srcSums[$column] }
                    $b = if ($tgtSums.IsNull($column)) { [decimal]0 } else { [decimal]$tgtSums[$column] }
                    if ($a -ne $b) { $sumMismatches++ }
                    $field = ($column.TrimStart('_') -replace '(Dt|Ct)$','')
                    $resourceResults.Add([pscustomobject]@{ register=$register.Value; table=$table; column=$column
                        field=$fieldNames[$field] + $(if ($column -match 'Dt$') { ' Дт' } elseif ($column -match 'Ct$') { ' Кт' } else { '' })
                        source_sum=$a; target_sum=$b; difference=$b-$a })
                }
            }

            # Accounting registers: OSV by account from movements.
            $accountMismatches = 0
            if ($type -match '^AccRg[0-9]+$' -and $names -contains '_AccountDtRRef') {
                $amount = @($numeric | Where-Object { $_ -notmatch '(Dt|Ct)$' -and $fieldNames[$_.TrimStart('_')] -eq 'Сумма' } | Select-Object -First 1)[0]
                if ($amount) {
                    # As 1C balances: only active movements, and the empty side
                    # of a one-sided (off-balance) entry is not an account.
                    $osvFilter = $filter + $(if ($names -contains '_Active') { ' AND s._Active=0x01' } else { '' })
                    $osv = {
                        param($Db)
@"
SELECT account, SUM(CASE WHEN s._Period<='$dbPeriodTo' THEN amount ELSE 0 END) AS balance,
       SUM(CASE WHEN s._Period>='$dbPeriodFrom' AND s._Period<='$dbPeriodTo' AND amount>0 THEN amount ELSE 0 END) AS turnover_dt,
       SUM(CASE WHEN s._Period>='$dbPeriodFrom' AND s._Period<='$dbPeriodTo' AND amount<0 THEN -amount ELSE 0 END) AS turnover_ct
FROM (SELECT s._Period, s._AccountDtRRef AS account, CONVERT(decimal(38,2),s.$(Quote $amount)) AS amount FROM $Db.dbo.$(Quote $table) s WHERE $osvFilter
      UNION ALL
      SELECT s._Period, s._AccountCtRRef, -CONVERT(decimal(38,2),s.$(Quote $amount)) FROM $Db.dbo.$(Quote $table) s WHERE $osvFilter) s
WHERE account<>0x00000000000000000000000000000000
GROUP BY account
"@
                    }
                    $accounts = Invoke-Table $connection @"
WITH a AS ($(& $osv $src)), b AS ($(& $osv $tgt))
SELECT COALESCE(acc._Code, sys.fn_varbintohexstr(COALESCE(a.account,b.account))) AS account_code,
       ISNULL(a.balance,0) AS source_balance, ISNULL(b.balance,0) AS target_balance,
       ISNULL(a.turnover_dt,0) AS source_turnover_dt, ISNULL(b.turnover_dt,0) AS target_turnover_dt,
       ISNULL(a.turnover_ct,0) AS source_turnover_ct, ISNULL(b.turnover_ct,0) AS target_turnover_ct
FROM a FULL JOIN b ON b.account=a.account
LEFT JOIN $src.dbo._Acc7 acc ON acc._IDRRef=COALESCE(a.account,b.account)
ORDER BY account_code;
"@
                    foreach ($row in $accounts.Rows) {
                        $diff = ([decimal]$row.target_balance -ne [decimal]$row.source_balance) -or
                                ([decimal]$row.target_turnover_dt -ne [decimal]$row.source_turnover_dt) -or
                                ([decimal]$row.target_turnover_ct -ne [decimal]$row.source_turnover_ct)
                        if ($diff) { $accountMismatches++ }
                        $accountResults.Add([pscustomobject]@{ register=$register.Value; table=$table; account=([string]$row.account_code).Trim()
                            source_balance=[decimal]$row.source_balance; target_balance=[decimal]$row.target_balance
                            source_turnover_dt=[decimal]$row.source_turnover_dt; target_turnover_dt=[decimal]$row.target_turnover_dt
                            source_turnover_ct=[decimal]$row.source_turnover_ct; target_turnover_ct=[decimal]$row.target_turnover_ct; match=(-not $diff) })
                    }
                }
            }

            $tableResults.Add([pscustomobject]@{ register=$register.Value; type=$type; table=$table; organization_column=$orgColumn
                source_rows=[int64]$c.source_rows; target_rows=[int64]$c.target_rows
                missing_in_target=$missing; extra_in_target=$extra; external_values=$externalValues
                resource_columns=$numeric.Count; resource_mismatches=$sumMismatches; account_mismatches=$accountMismatches
                result=if ($missing -eq 0 -and $extra -eq 0 -and $sumMismatches -eq 0 -and $accountMismatches -eq 0) { 'PASS' }
                       elseif ($exceptions.ContainsKey([string]$register.Value)) { 'EXPECTED' } else { 'FAIL' }
                exception=$exceptions[[string]$register.Value] })
            Write-Host ("{0} {1}: source {2}, target {3}, missing {4}, extra {5}" -f $register.Value, $table, $c.source_rows, $c.target_rows, $missing, $extra)

            # Rows that belong to the compared movements by recorder and line:
            # extra dimensions of accounting movements, actual action period of
            # calculation movements.
            $linkedPattern = if ($type -match '^AccRg[0-9]+$') { '^AccRgED[0-9]+$' } elseif ($type -match '^CRg[0-9]+$') { '^CRgActP[0-9]+$' } else { $null }
            $isCalc = $type -match '^CRg[0-9]+$'
            $linkedLabel = if ($isCalc) { ' (период действия)' } else { ' (субконто)' }
            $linkedHostLabel = if ($isCalc) { $linkedLabel } else { ' ED' }
            if ($linkedPattern) {
                foreach ($edType in @($typeNames.GetEnumerator() | Where-Object { $_.Value -eq $register.Value -and $_.Name -match $linkedPattern } | ForEach-Object Name)) {
                    $suffix = [regex]::Match($table, 'X[0-9]+$').Value
                    $edTable = "_$edType$suffix"
                    $edExists = Invoke-Table $connection "SELECT COUNT(*) AS n FROM $tgt.sys.tables WHERE name=N'$edTable';"
                    if ([int]$edExists.Rows[0].n -eq 0) { continue }
                    # Calculation recorders are composite; accounting keeps its original join.
                    $edTypeMatch = if ($isCalc) { 's._RecorderTRef=e._RecorderTRef AND ' } else { '' }
                    $edFilter = { param($Db) "EXISTS(SELECT 1 FROM $Db.dbo.$(Quote $table) s WHERE $edTypeMatch" + "s._RecorderRRef=e._RecorderRRef AND s._LineNo=e._LineNo AND $filter)" }
                    $ed = Invoke-Table $connection @"
SELECT BINARY_CHECKSUM(*) AS h INTO #s FROM $src.dbo.$(Quote $edTable) e WHERE $(& $edFilter $src);
SELECT BINARY_CHECKSUM(*) AS h INTO #t FROM $tgt.dbo.$(Quote $edTable) e WHERE $(& $edFilter $tgt);
WITH a AS (SELECT h, COUNT_BIG(*) AS c FROM #s GROUP BY h), b AS (SELECT h, COUNT_BIG(*) AS c FROM #t GROUP BY h)
SELECT (SELECT COUNT_BIG(*) FROM #s) AS source_rows, (SELECT COUNT_BIG(*) FROM #t) AS target_rows,
       ISNULL(SUM(CASE WHEN ISNULL(a.c,0)>ISNULL(b.c,0) THEN ISNULL(a.c,0)-ISNULL(b.c,0) ELSE 0 END),0) AS missing_in_target,
       ISNULL(SUM(CASE WHEN ISNULL(b.c,0)>ISNULL(a.c,0) THEN ISNULL(b.c,0)-ISNULL(a.c,0) ELSE 0 END),0) AS extra_in_target
FROM a FULL JOIN b ON b.h=a.h;
DROP TABLE #s; DROP TABLE #t;
"@
                    $e = $ed.Rows[0]
                    $tableResults.Add([pscustomobject]@{ register=$register.Value + $linkedLabel; type=$edType; table=$edTable; organization_column='via movements'
                        source_rows=[int64]$e.source_rows; target_rows=[int64]$e.target_rows
                        missing_in_target=[int64]$e.missing_in_target; extra_in_target=[int64]$e.extra_in_target; external_values=[int64]0
                        resource_columns=0; resource_mismatches=0; account_mismatches=0
                        result=if ([int64]$e.missing_in_target -eq 0 -and [int64]$e.extra_in_target -eq 0) { 'PASS' } else { 'FAIL' }
                        exception=$null })
                    Write-Host ("{0} {1}: source {2}, target {3}, missing {4}, extra {5}" -f ($register.Value + $linkedHostLabel), $edTable, $e.source_rows, $e.target_rows, $e.missing_in_target, $e.extra_in_target)
                }
            }
        }
    }

    $tablesPath = Join-Path $OutputDirectory 'reconcile-tables.csv'
    $resourcesPath = Join-Path $OutputDirectory 'reconcile-resources.csv'
    $accountsPath = Join-Path $OutputDirectory 'reconcile-accounts.csv'
    # Decimals are written in the invariant culture so the CSV round-trips
    # regardless of the OS locale (a ru-RU comma was read back as a thousands
    # separator).  The objects themselves stay numeric for the summary below.
    $inv = [Globalization.CultureInfo]::InvariantCulture
    function ConvertTo-InvariantRow($Row) {
        $copy = [ordered]@{}
        foreach ($p in $Row.PSObject.Properties) { $copy[$p.Name] = if ($p.Value -is [decimal]) { $p.Value.ToString($inv) } else { $p.Value } }
        [pscustomobject]$copy
    }
    $tableResults | Export-Csv -LiteralPath $tablesPath -NoTypeInformation -Encoding UTF8
    $resourceResults | ForEach-Object { ConvertTo-InvariantRow $_ } | Export-Csv -LiteralPath $resourcesPath -NoTypeInformation -Encoding UTF8
    $accountResults | ForEach-Object { ConvertTo-InvariantRow $_ } | Export-Csv -LiteralPath $accountsPath -NoTypeInformation -Encoding UTF8
    $failed = @($tableResults | Where-Object result -eq 'FAIL')
    $expected = @($tableResults | Where-Object result -eq 'EXPECTED')
    $summary = [ordered]@{
        generated_at_local=(Get-Date).ToString('s'); server=$Server; read_only=$true
        source_database=$SourceDatabase; target_database=$TargetDatabase
        organization_code=$organizationFilter.label; organization_id_hex=$orgId
        period_from=$PeriodFrom.ToString('yyyy-MM-dd'); period_to=$PeriodTo.ToString('yyyy-MM-dd')
        registers=@($registers | ForEach-Object Value); tables=$tableResults.Count
        source_rows=[int64](($tableResults | Measure-Object source_rows -Sum).Sum)
        target_rows=[int64](($tableResults | Measure-Object target_rows -Sum).Sum)
        missing_in_target=[int64](($tableResults | Measure-Object missing_in_target -Sum).Sum)
        extra_in_target=[int64](($tableResults | Measure-Object extra_in_target -Sum).Sum)
        resource_mismatches=@($resourceResults | Where-Object difference -ne 0).Count
        accounts=$accountResults.Count; account_mismatches=@($accountResults | Where-Object { -not $_.match }).Count
        failed_tables=@($failed | ForEach-Object { "$($_.register) $($_.table)" })
        expected_tables=@($expected | ForEach-Object { "$($_.register) $($_.table): $($_.exception)" })
        result=if ($failed.Count -eq 0) { 'PASS' } else { 'FAIL' }
        tables_csv=$tablesPath; resources_csv=$resourcesPath; accounts_csv=$accountsPath
        duration_seconds=[Math]::Round(((Get-Date)-$started).TotalSeconds,1)
    }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
    $summary | ConvertTo-Json -Depth 5
}
finally { $connection.Dispose() }
