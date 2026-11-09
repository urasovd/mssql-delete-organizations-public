[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$Database,
    [Parameter(Mandatory)] [string]$ObjectManifest,
    [Parameter(Mandatory)] [string]$FixedEdgesManifest,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-batched-delete')),
    [ValidateRange(100,50000)] [int]$BatchSize = 50000,
    [ValidateRange(100,50000)] [int]$RegisterBatchSize = 50000,
    [ValidateRange(100,50000)] [int]$JournalBatchSize = 50000,
    [ValidateRange(100,50000)] [int]$TabularBatchSize = 50000,
    [ValidateRange(100,50000)] [int]$ObjectBatchSize = 50000,
    [ValidateSet('Organizations','Documents','Mixed')]
    [string]$RootMode = 'Organizations',
    [string]$ExpectedManifestSha256,
    # Kept organizations by reference (see lib/Organizations.ps1), one or more,
    # or by code where the catalog has one.
    [Alias('PreservedOrganizationId')] [string[]]$PreservedOrganization,
    [string[]]$PreservedOrganizationCode,
    # Physical type of the Организации catalog; by default read from
    # type-names.csv next to -FixedEdgesManifest (Reference53 in BIT,
    # Reference132 in ZUP).
    [ValidatePattern('^(Reference[0-9]+)?$')] [string]$OrganizationTypeTable,
    [string]$ClearReferenceTabularTablesManifest,
    [string]$ExpectedClearReferenceTabularTablesSha256,
    [string]$ProtectedRegisterRowsManifest,
    [string]$ExpectedProtectedRegisterRowsSha256,
    [switch]$BoundaryOnly,
    [string]$SimulatedRuleActionsPlan,
    [switch]$Apply,
    [string]$ConfirmationToken
)

<#
    Batched physical delete for one frozen object manifest.

    Scope:
      * dependent document rows, recorders, journals and tabular parts of all
        object types (documents, catalogs/references, characteristics,
        business processes, tasks and exchange plans);
      * all non-derived rows in information registers, movements and sequences
        that contain a fixed or composite reference to the target object-set;
      * primary object rows from the manifest.

    Intentionally excluded:
      * accumulation/accounting totals and other derived tables;
      * _InfoRgSL* slices;
      * every row is deleted only in short committed batches.

    The script is fail-closed: without -Apply it is read-only; with -Apply it
    requires the exact confirmation token. The target database is never altered
    by this script's metadata discovery, only by the explicit DELETE statements.
#>

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\Organizations.ps1')
$preservedFilter = Get-OrganizationFilter $PreservedOrganization $PreservedOrganizationCode
if (-not $OrganizationTypeTable) {
    $typeNamesPath = Join-Path (Split-Path -Parent $FixedEdgesManifest) 'type-names.csv'
    if (-not (Test-Path -LiteralPath $typeNamesPath -PathType Leaf)) { throw "Pass -OrganizationTypeTable: $typeNamesPath not found." }
    $orgTypes = @(Import-Csv -LiteralPath $typeNamesPath | Where-Object { $_.logical_name -eq 'Организации' -and $_.type_table -match '^Reference[0-9]+$' })
    if ($orgTypes.Count -ne 1) { throw "Expected one Организации catalog in $typeNamesPath; found $($orgTypes.Count)." }
    $OrganizationTypeTable = [string]$orgTypes[0].type_table
}
$organizationTypeCode = [int]($OrganizationTypeTable -replace '^Reference','')
$expectedToken = switch ($RootMode) {
    'Documents' { 'DELETE-MARKED-DOCUMENTS' }
    'Mixed' { 'DELETE-MIXED-MANIFEST' }
    default { 'DELETE-ALL-MARKED' }
}
if ($Apply -and $ConfirmationToken -ne $expectedToken -and -not ($RootMode -eq 'Organizations' -and $ConfirmationToken -eq 'DELETE-EKATERINBURG')) {
    throw "RootMode=$RootMode requires confirmation token $expectedToken."
}
if ($Apply -and $BoundaryOnly) { throw '-BoundaryOnly is read-only and cannot be combined with -Apply.' }
# A read-only preview before script 42 has run: conflicts in tables that the
# rule-action plan of script 42 will change are reported as covered instead
# of blocking, so the dry-run can be counted.  Never valid with -Apply: the
# real run must see PASS after script 42 was applied.
if ($Apply -and $SimulatedRuleActionsPlan) { throw '-SimulatedRuleActionsPlan is read-only and cannot be combined with -Apply.' }
$simulatedRuleRows = @{}
if ($SimulatedRuleActionsPlan) {
    foreach ($row in Import-Csv -LiteralPath $SimulatedRuleActionsPlan) {
        $simulatedRuleRows[[string]$row.table] = [long]$simulatedRuleRows[[string]$row.table] + [long]$row.rows
    }
}
if ($Apply -and $RootMode -eq 'Organizations' -and -not $ProtectedRegisterRowsManifest) {
    throw '-Apply with RootMode=Organizations requires -ProtectedRegisterRowsManifest.'
}
if ($Apply) {
    if ($ExpectedManifestSha256 -notmatch '^[0-9A-Fa-f]{64}$') {
        throw '-Apply requires -ExpectedManifestSha256 for the frozen object manifest.'
    }
    $actualManifestSha256 = (Get-FileHash -LiteralPath $ObjectManifest -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualManifestSha256 -ne $ExpectedManifestSha256.ToLowerInvariant()) {
        throw "Object manifest SHA256 changed: $actualManifestSha256. Expected: $($ExpectedManifestSha256.ToLowerInvariant())."
    }
} else {
    $actualManifestSha256 = (Get-FileHash -LiteralPath $ObjectManifest -Algorithm SHA256).Hash.ToLowerInvariant()
}

$clearReferenceTabularTables = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$clearReferenceTabularTablesSha256 = $null
if ($ClearReferenceTabularTablesManifest) {
    if (-not (Test-Path -LiteralPath $ClearReferenceTabularTablesManifest -PathType Leaf)) {
        throw "CLEAR table manifest was not found: $ClearReferenceTabularTablesManifest"
    }
    if ($ExpectedClearReferenceTabularTablesSha256 -notmatch '^[0-9A-Fa-f]{64}$') {
        throw '-ClearReferenceTabularTablesManifest requires -ExpectedClearReferenceTabularTablesSha256.'
    }
    $clearReferenceTabularTablesSha256 = (Get-FileHash -LiteralPath $ClearReferenceTabularTablesManifest -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($clearReferenceTabularTablesSha256 -ne $ExpectedClearReferenceTabularTablesSha256.ToLowerInvariant()) {
        throw "CLEAR table manifest SHA256 changed: $clearReferenceTabularTablesSha256."
    }
    foreach ($row in Import-Csv -LiteralPath $ClearReferenceTabularTablesManifest) {
        $table = [string]$row.physical_table
        if ($table -notmatch '^_Reference[0-9]+_VT[0-9]+(?:X[0-9]+)?$') {
            throw "CLEAR policy is restricted to Reference tabular tables; invalid table: $table"
        }
        [void]$clearReferenceTabularTables.Add($table)
    }
    if ($clearReferenceTabularTables.Count -eq 0) { throw 'CLEAR table manifest is empty.' }
}

$protectedRegisterRowsSha256 = $null
$protectedRegisterLocations = @{}
if ($ProtectedRegisterRowsManifest) {
    if (-not (Test-Path -LiteralPath $ProtectedRegisterRowsManifest -PathType Leaf)) {
        throw "Protected register-row manifest was not found: $ProtectedRegisterRowsManifest"
    }
    if ($ExpectedProtectedRegisterRowsSha256 -notmatch '^[0-9A-Fa-f]{64}$') {
        throw '-ProtectedRegisterRowsManifest requires -ExpectedProtectedRegisterRowsSha256.'
    }
    $protectedRegisterRowsSha256 = (Get-FileHash -LiteralPath $ProtectedRegisterRowsManifest -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($protectedRegisterRowsSha256 -ne $ExpectedProtectedRegisterRowsSha256.ToLowerInvariant()) {
        throw "Protected register-row manifest SHA256 changed: $protectedRegisterRowsSha256."
    }
    foreach ($row in Import-Csv -LiteralPath $ProtectedRegisterRowsManifest) {
        if ([int64]$row.matched_rows -le 0) { continue }
        $table = [string]$row.physical_table
        $idColumn = [string]$row.id_column
        $typeColumn = [string]$row.type_column
        if ($table -notmatch '^_[A-Za-z0-9_]+$' -or $idColumn -notmatch '^_[A-Za-z0-9_]+$') {
            throw "Unsafe protected register location: $table/$idColumn"
        }
        if ($typeColumn -and $typeColumn -notmatch '^_[A-Za-z0-9_]+$') {
            throw "Unsafe protected register type column: $typeColumn"
        }
        $key = "$typeColumn|$idColumn"
        if (-not $protectedRegisterLocations.ContainsKey($table)) {
            $protectedRegisterLocations[$table] = [ordered]@{}
        }
        $protectedRegisterLocations[$table][$key] = [pscustomobject]@{ type_column=$typeColumn; id_column=$idColumn }
    }
    if ($protectedRegisterLocations.Count -eq 0) { throw 'Protected register-row manifest contains no positive locations.' }
}

function Quote-SqlIdentifier([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -notmatch '^_[A-Za-z0-9_]+$') {
        throw "Unsafe SQL identifier: $Name"
    }
    return '[' + $Name + ']'
}

function Assert-TypeToken([string]$Name) {
    if ($Name -notmatch '^[A-Za-z]+[0-9]+$') { throw "Unsafe 1C type token: $Name" }
}

function Convert-HexToBytes16([string]$Hex) {
    if ($Hex -notmatch '^0x[0-9a-fA-F]{32}$') { throw "Expected binary(16) hex value, got '$Hex'." }
    $result = [byte[]]::new(16)
    for ($index = 0; $index -lt 16; $index++) {
        $result[$index] = [Convert]::ToByte($Hex.Substring(2 + ($index * 2), 2), 16)
    }
    Write-Output -NoEnumerate $result
}

function Invoke-Scalar([System.Data.SqlClient.SqlConnection]$Connection, [System.Data.SqlClient.SqlTransaction]$Transaction, [string]$Sql, [int]$TimeoutSeconds = 900) {
    $command = $Connection.CreateCommand()
    $command.Transaction = $Transaction
    $command.CommandTimeout = $TimeoutSeconds
    $command.CommandText = "SET NOCOUNT ON;`n$Sql"
    try { return $command.ExecuteScalar() } finally { $command.Dispose() }
}

function Invoke-Table([System.Data.SqlClient.SqlConnection]$Connection, [System.Data.SqlClient.SqlTransaction]$Transaction, [string]$Sql, [int]$TimeoutSeconds = 900) {
    $command = $Connection.CreateCommand()
    $command.Transaction = $Transaction
    $command.CommandTimeout = $TimeoutSeconds
    $command.CommandText = "SET NOCOUNT ON;`n$Sql"
    try {
        $adapter = [System.Data.SqlClient.SqlDataAdapter]::new($command)
        $adapter.SelectCommand.Transaction = $Transaction
        $table = [System.Data.DataTable]::new()
        [void]$adapter.Fill($table)
        Write-Output -NoEnumerate $table
    }
    finally { $command.Dispose() }
}

function Is-DerivedTable([string]$Name) {
    # Platform-maintained slices, totals, aggregates and calculation buffers
    # are deliberately left for a later rebuild by 1C.  _AccRgED is not in
    # this list: it stores the extra-dimension values of accounting-register
    # records and therefore has to be deleted together with those records.
    return ($Name -match '^_InfoRgSL' -or
            $Name -match '^_AccumRg(?:T|AT|Agg|St|Bf|Dl)' -or
            $Name -match '^_AccumRgOpt(?:$|X[0-9]+$)' -or
            $Name -match '^_AccRg(?:T|AT|CT)' -or
            $Name -match '^_AccRgOpt(?:$|X[0-9]+$)')
}

function Get-SourceRole([string]$Name) {
    if ($Name -match '^_InfoRg') { return 'information-register' }
    if ($Name -match '^_Seq') { return 'sequence' }
    if ($Name -match '^_AccumRg[0-9]') { return 'accumulation-movement' }
    if ($Name -match '^_AccRg[0-9]') { return 'accounting-movement' }
    if ($Name -match '^_DocumentJournal') { return 'document-journal' }
    if ($Name -match '^_[A-Za-z]+[0-9]+(?:X[0-9]+)?$') { return 'object' }
    return 'other'
}

function Add-DeletePlanCondition(
    [hashtable]$Plans,
    [string]$Table,
    [string]$Condition,
    [string]$Reason,
    [string]$Role,
    [int]$Phase
) {
    if (-not $Plans.ContainsKey($Table)) {
        $Plans[$Table] = [pscustomobject]@{
            table = $Table
            phase = $Phase
            conditions = [System.Collections.Generic.List[string]]::new()
            reasons = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            roles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        }
    }
    $plan = $Plans[$Table]
    if ($Phase -gt $plan.phase) { $plan.phase = $Phase }
    if (-not $plan.conditions.Contains($Condition)) { $plan.conditions.Add($Condition) }
    [void]$plan.reasons.Add($Reason)
    [void]$plan.roles.Add($Role)
}

function Add-BoundaryReferenceCondition(
    [hashtable]$BoundaryPlans,
    [string]$Table,
    [string]$Condition,
    [string]$ReferencePath,
    [string]$TargetJoinPredicate
) {
    if (-not $BoundaryPlans.ContainsKey($Table)) {
        $BoundaryPlans[$Table] = [pscustomobject]@{
            table = $Table
            conditions = [System.Collections.Generic.List[string]]::new()
            reference_paths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            references = [System.Collections.Generic.List[object]]::new()
        }
    }
    $plan = $BoundaryPlans[$Table]
    if (-not $plan.conditions.Contains($Condition)) { $plan.conditions.Add($Condition) }
    [void]$plan.reference_paths.Add($ReferencePath)
    [void]$plan.references.Add([pscustomobject]@{ path=$ReferencePath; target_join=$TargetJoinPredicate })
}

function Add-OwnerDescriptor(
    [hashtable]$OwnerDescriptors,
    [string]$Table,
    [string]$TypeTableExpression,
    [string]$TypeCodeExpression,
    [string]$IdExpression,
    [string]$Role
) {
    if (-not $OwnerDescriptors.ContainsKey($Table)) {
        $OwnerDescriptors[$Table] = [System.Collections.Generic.List[object]]::new()
    }
    [void]$OwnerDescriptors[$Table].Add([pscustomobject]@{
        type_table_expression=$TypeTableExpression
        type_code_expression=$TypeCodeExpression
        id_expression=$IdExpression
        role=$Role
    })
}

function Get-DeletePlanBatchSize(
    [object]$Plan,
    [int]$DefaultBatchSize,
    [int]$RegisterBatchSize,
    [int]$JournalBatchSize,
    [int]$TabularBatchSize,
    [int]$ObjectBatchSize
) {
    if ($Plan.roles.Contains('object')) { return $ObjectBatchSize }
    if ($Plan.roles.Contains('tabular')) { return $TabularBatchSize }
    if ($Plan.roles.Contains('document-journal')) { return $JournalBatchSize }
    if ($Plan.roles.Contains('recorder') -or
        $Plan.roles.Contains('information-register') -or
        $Plan.roles.Contains('accumulation-movement') -or
        $Plan.roles.Contains('accounting-movement')) { return $RegisterBatchSize }
    return $DefaultBatchSize
}

function Get-DeletePlanPredicate(
    [object]$Plan,
    [hashtable]$ProtectedPredicates
) {
    $predicate = '(' + (@($Plan.conditions) -join ' OR ') + ')'
    if ($ProtectedPredicates.ContainsKey([string]$Plan.table) -and $ProtectedPredicates[[string]$Plan.table].Count -gt 0) {
        # Keep the whole owner-delete expression atomic because the boundary
        # audit embeds it below NOT.  Without the outer parentheses SQL parses
        # `NOT (owner) AND NOT (protected)` and hides protected conflicts.
        $predicate = '(' + $predicate + ' AND NOT (' + (@($ProtectedPredicates[[string]$Plan.table]) -join ' OR ') + '))'
    }
    return $predicate
}

function Import-TargetManifest([System.Data.SqlClient.SqlConnection]$Connection, [string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Manifest not found: $Path" }
    $table = [System.Data.DataTable]::new()
    [void]$table.Columns.Add('type_table', [string])
    [void]$table.Columns.Add('type_code', [int])
    [void]$table.Columns.Add('id', [byte[]])
    $reader = [System.IO.StreamReader]::new((Resolve-Path -LiteralPath $Path))
    try {
        $header = $reader.ReadLine()
        if ($header -ne '"type_table","type_code","id_hex","graph_depth","inclusion_reason"') { throw 'Unexpected manifest header.' }
        while (($line = $reader.ReadLine()) -ne $null) {
            $parts = $line.Trim('"').Split('" ,"'.Replace(' ',''))
            if ($parts.Count -ne 5) { continue }
            Assert-TypeToken $parts[0]
            $row = $table.NewRow()
            $row['type_table'] = $parts[0]
            $row['type_code'] = [int]$parts[1]
            $row['id'] = [byte[]](Convert-HexToBytes16 $parts[2])
            $table.Rows.Add($row)
        }
    }
    finally { $reader.Dispose() }
    if ($table.Rows.Count -eq 0) { throw 'Manifest contains no rows.' }

    $bulk = [System.Data.SqlClient.SqlBulkCopy]::new($Connection)
    try {
        $bulk.DestinationTableName = '#TargetObjects'
        $bulk.BatchSize = 10000
        $bulk.BulkCopyTimeout = 900
        [void]$bulk.ColumnMappings.Add('type_table','type_table')
        [void]$bulk.ColumnMappings.Add('type_code','type_code')
        [void]$bulk.ColumnMappings.Add('id','id')
        $bulk.WriteToServer($table)
    }
    finally { $bulk.Dispose() }
    return $table.Rows.Count
}

function Add-BatchLog([System.Collections.Generic.List[object]]$Log, [string]$Action, [int]$BatchNumber, [long]$Rows, [double]$Seconds) {
    [void]$Log.Add([pscustomobject]@{
        recorded_at = (Get-Date).ToString('s')
        action = $Action
        batch_number = $BatchNumber
        deleted_rows = $Rows
        duration_seconds = [math]::Round($Seconds, 3)
    })
}

function Invoke-BatchedDelete(
    [System.Data.SqlClient.SqlConnection]$Connection,
    [string]$Action,
    [string]$DeleteSql,
    [System.Collections.Generic.List[object]]$Log,
    [int]$BatchSize,
    [switch]$Apply,
    [int]$TimeoutSeconds = 900
) {
    $total = [long]0
    $batchNumber = 0
    while ($true) {
        $batchNumber++
        $started = Get-Date
        if (-not $Apply) {
            $countSql = "SELECT COUNT_BIG(*) FROM dbo.$DeleteSql;"
            $count = [long](Invoke-Scalar $Connection $null $countSql $TimeoutSeconds)
            Add-BatchLog $Log $Action $batchNumber $count ((Get-Date)-$started).TotalSeconds
            return $count
        }

        $transaction = $Connection.BeginTransaction([System.Data.IsolationLevel]::ReadCommitted)
        try {
            $command = $Connection.CreateCommand()
            $command.Transaction = $transaction
            $command.CommandTimeout = $TimeoutSeconds
            $command.CommandText = "SET XACT_ABORT ON; SET NOCOUNT ON; DELETE TOP (@BatchSize) src FROM dbo.$DeleteSql; SELECT CONVERT(bigint,@@ROWCOUNT);"
            [void]$command.Parameters.Add('@BatchSize',[System.Data.SqlDbType]::Int)
            $command.Parameters['@BatchSize'].Value = $BatchSize
            $rows = [long]$command.ExecuteScalar()
            $command.Dispose()
            $transaction.Commit()
        }
        catch {
            try { $transaction.Rollback() } catch {}
            throw "Action '$Action' failed in batch ${batchNumber}: $($_.Exception.Message)"
        }
        finally { if ($null -ne $transaction) { $transaction.Dispose() } }

        Add-BatchLog $Log $Action $batchNumber $rows ((Get-Date)-$started).TotalSeconds
        $total += $rows
        if ($rows -eq 0) { return $total }
    }
}

$connectionString = "Server=$Server;Database=$Database;Integrated Security=True;Encrypt=False;Application Name=Codex-1C-BatchedDelete"
$connection = [System.Data.SqlClient.SqlConnection]::new($connectionString)
$log = [System.Collections.Generic.List[object]]::new()
$startedOverall = Get-Date

try {
    $connection.Open()
    New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    $OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
    $identity = Invoke-Table $connection $null "SELECT @@SERVERNAME AS server_name, DB_NAME() AS database_name, d.state_desc, d.user_access_desc, d.recovery_model_desc FROM sys.databases d WHERE d.database_id=DB_ID();" 30
    if ($identity.Rows.Count -ne 1 -or [string]$identity.Rows[0].database_name -ne $Database) { throw "Unexpected target database. Expected $Database." }
    if ([string]$identity.Rows[0].state_desc -ne 'ONLINE') { throw 'Target database is not ONLINE.' }

    [void](Invoke-Scalar $connection $null @'
CREATE TABLE #TargetObjects(type_table nvarchar(128) NOT NULL,type_code int NOT NULL,id binary(16) NOT NULL,PRIMARY KEY(type_table,id));
CREATE TABLE #TargetDocuments(type_code int NOT NULL,type_ref binary(4) NOT NULL,id binary(16) NOT NULL,PRIMARY KEY(type_code,id));
CREATE INDEX IX_TargetObjects_TypeCodeId ON #TargetObjects(type_code,id);
CREATE INDEX IX_TargetDocuments_TypeRefId ON #TargetDocuments(type_ref,id);
CREATE TABLE #KeepBoundaryEvidence
(
    type_table nvarchar(128) NOT NULL,
    type_code int NOT NULL,
    id binary(16) NOT NULL,
    source_table nvarchar(128) NOT NULL,
    reference_path nvarchar(512) NOT NULL,
    owner_type_table nvarchar(128) NOT NULL,
    owner_type_code int NOT NULL,
    owner_id binary(16) NOT NULL,
    owner_role nvarchar(64) NOT NULL
);
'@ 60)
    $loaded = Import-TargetManifest $connection $ObjectManifest
    [void](Invoke-Scalar $connection $null "INSERT #TargetDocuments(type_code,type_ref,id) SELECT type_code,CONVERT(binary(4),type_code),id FROM #TargetObjects WHERE type_table LIKE N'Document%'; SELECT 1;" 60)
    $documentCount = [long](Invoke-Scalar $connection $null 'SELECT COUNT_BIG(*) FROM #TargetDocuments;' 60)
    $organizationCount = [long](Invoke-Scalar $connection $null "SELECT COUNT_BIG(*) FROM #TargetObjects WHERE type_table=N'$OrganizationTypeTable';" 60)
    switch ($RootMode) {
        'Organizations' { if ($organizationCount -lt 1) { throw 'RootMode=Organizations requires at least one target organization.' } }
        'Documents' {
            if ($organizationCount -ne 0) { throw "RootMode=Documents forbids target organizations; found $organizationCount." }
            if ($documentCount -lt 1) { throw 'RootMode=Documents requires at least one target document.' }
        }
        'Mixed' { if ($loaded -lt 1) { throw 'RootMode=Mixed requires a non-empty manifest.' } }
    }
    $preservedCommand = $connection.CreateCommand()
    $preservedCommand.CommandTimeout = 60
    $preservedCommand.CommandText = @"
SET NOCOUNT ON;
SELECT COUNT_BIG(*)
FROM dbo._$OrganizationTypeTable organization
JOIN #TargetObjects target ON target.type_table=N'$OrganizationTypeTable' AND target.id=organization._IDRRef
WHERE $((Get-OrganizationFilter $PreservedOrganization $PreservedOrganizationCode 'organization').sql);
"@
    try { $preservedInManifest = [long]$preservedCommand.ExecuteScalar() }
    finally { $preservedCommand.Dispose() }
    if ($preservedInManifest -ne 0) { throw "Preserved organization '$($preservedFilter.label)' is present in the frozen manifest." }

    $preservedRows = Invoke-Table $connection $null @"
SELECT sys.fn_varbintohexstr(_IDRRef) AS id_hex,
       CONVERT(int,CONVERT(varbinary(1),_Marked)) AS marked
FROM dbo._$OrganizationTypeTable WITH(READUNCOMMITTED)
WHERE $($preservedFilter.sql);
"@ 60
    if ($preservedRows.Rows.Count -ne $preservedFilter.expected) {
        throw "Expected $($preservedFilter.expected) preserved organization(s) '$($preservedFilter.label)'; found $($preservedRows.Rows.Count)."
    }
    foreach ($row in $preservedRows.Rows) { if ([int]$row.marked -ne 0) { throw "Preserved organization $($row.id_hex) is marked." } }
    $preservedIds = @($preservedRows.Rows | ForEach-Object { [string]$_.id_hex } | Sort-Object)
    $preservedOrganizationIdHex = $preservedIds -join ','
    $preservedIdList = $preservedIds -join ','

    $protectedRegisterPredicates = @{}
    foreach ($table in $protectedRegisterLocations.Keys) {
        $predicates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($location in $protectedRegisterLocations[$table].Values) {
            $id = Quote-SqlIdentifier ([string]$location.id_column)
            if ([string]::IsNullOrWhiteSpace([string]$location.type_column)) {
                [void]$predicates.Add("src.$id IN ($preservedIdList)")
            }
            else {
                $type = Quote-SqlIdentifier ([string]$location.type_column)
                [void]$predicates.Add("(CONVERT(int,CONVERT(varbinary(4),src.$type))=$organizationTypeCode AND src.$id IN ($preservedIdList))")
            }
        }
        $protectedRegisterPredicates[$table] = $predicates
    }

    $fixedEdges = @()
    if (Test-Path -LiteralPath $FixedEdgesManifest -PathType Leaf) {
        if ([System.IO.Path]::GetFileName($FixedEdgesManifest) -eq 'level1-fixed-edges.csv') {
            throw 'level1-fixed-edges.csv is a closure artifact and excludes document journals. Pass final-fixed-reference-inventory/fixed-reference-edges.csv from script 12.'
        }
        $fixedEdges = @(Import-Csv -LiteralPath $FixedEdgesManifest | Where-Object {
            $_.physical_table -match '^_[A-Za-z0-9_]+$' -and
            $_.physical_column -match '^_[A-Za-z0-9_]+$' -and
            $_.target_table -match '^[A-Za-z]+[0-9]+$' -and
            -not (Is-DerivedTable ([string]$_.physical_table))
        })
    }

    $pairTable = Invoke-Table $connection $null @'
SELECT t.name AS physical_table,type_ref.name AS type_column,id_ref.name AS id_column
FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id AND s.name=N'dbo'
JOIN sys.columns type_ref ON type_ref.object_id=t.object_id AND type_ref.name LIKE N'%TRef' AND type_ref.max_length=4
JOIN sys.columns id_ref ON id_ref.object_id=t.object_id AND id_ref.name=LEFT(type_ref.name,LEN(type_ref.name)-4)+N'RRef' AND id_ref.max_length=16
WHERE t.is_ms_shipped=0 AND t.name NOT LIKE N'[_]InfoRgSL%'
ORDER BY t.name,type_ref.column_id;
'@ 120
    $pairs = @($pairTable.Rows | ForEach-Object {
        $name = [string]$_.physical_table
        if ($name -match '^_[A-Za-z0-9_]+$' -and -not (Is-DerivedTable $name) -and
            [string]$_.type_column -notin @('_RecorderTRef','_Recorder_RTRef','_DocumentTRef','_Document_RTRef')) {
            [pscustomobject]@{ physical_table=$name; type_column=[string]$_.type_column; id_column=[string]$_.id_column }
        }
    })

    # Atomic storage must never be filtered by an arbitrary reference stored
    # inside a row.  A document movement, tabular row, journal row or object
    # row is deleted only together with its owner.  Applying fixed/composite
    # predicates to these tables can otherwise delete only some movements of
    # a preserved document (the defect found after the v5b run).
    $atomicTableRows = Invoke-Table $connection $null @'
SELECT DISTINCT t.name AS table_name
FROM sys.tables t
WHERE t.is_ms_shipped=0
  AND
  (
      (EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_IDRRef' AND c.max_length=16)
       AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_Marked' AND c.max_length=1))
   OR (EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_RecorderTRef' AND c.max_length=4)
       AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_RecorderRRef' AND c.max_length=16))
   OR (EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_Recorder_RTRef' AND c.max_length=4)
       AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_Recorder_RRRef' AND c.max_length=16))
   OR (EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_DocumentTRef' AND c.max_length=4)
       AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_DocumentRRef' AND c.max_length=16))
   OR (t.name LIKE N'[_]%[_]VT%'
       AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name LIKE N'[_]%[_]IDRRef' AND c.max_length=16))
  );
'@ 120
    $atomicTables = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $atomicTableRows.Rows) { [void]$atomicTables.Add([string]$row.table_name) }

    # Build one predicate per physical table before executing DML.  The legacy
    # executor made separate linked-object and recorder passes over the same
    # large register.  On _AccRgED7952X1 alone that duplicated scan accounted
    # for roughly 77% of a 2h16m run.  A table now receives one DELETE loop
    # containing every applicable reason.
    $plans = @{}
    $boundaryPlans = @{}
    $ownerDescriptors = @{}
    foreach ($edge in $fixedEdges) {
        $table = [string]$edge.physical_table
        $targetJoin = "target.type_table=N'$($edge.target_table)' AND target.id=src.$(Quote-SqlIdentifier ([string]$edge.physical_column))"
        $condition = "EXISTS (SELECT 1 FROM #TargetObjects target WHERE $targetJoin)"
        if ($atomicTables.Contains($table)) {
            if ($clearReferenceTabularTables.Contains($table)) {
                Add-DeletePlanCondition $plans $table $condition 'explicit-clear-fixed-reference' 'tabular' 20
                continue
            }
            Add-BoundaryReferenceCondition $boundaryPlans $table $condition ("fixed:{0}->{1}" -f [string]$edge.physical_column,[string]$edge.target_table) $targetJoin
            continue
        }
        Add-DeletePlanCondition $plans $table $condition 'fixed-reference' (Get-SourceRole $table) 20
    }
    foreach ($pair in $pairs) {
        $table = [string]$pair.physical_table
        $targetJoin = "target.type_code=CONVERT(int,CONVERT(varbinary(4),src.$(Quote-SqlIdentifier $pair.type_column))) AND target.id=src.$(Quote-SqlIdentifier $pair.id_column)"
        $condition = "EXISTS (SELECT 1 FROM #TargetObjects target WHERE $targetJoin)"
        if ($atomicTables.Contains($table)) {
            if ($clearReferenceTabularTables.Contains($table)) {
                Add-DeletePlanCondition $plans $table $condition 'explicit-clear-composite-reference' 'tabular' 20
                continue
            }
            Add-BoundaryReferenceCondition $boundaryPlans $table $condition ("composite:{0}+{1}" -f $pair.type_column,$pair.id_column) $targetJoin
            continue
        }
        Add-DeletePlanCondition $plans $table $condition 'composite-reference' (Get-SourceRole $table) 20
    }

    $recorderTables = Invoke-Table $connection $null @'
SELECT t.name AS table_name,N'_RecorderTRef' AS type_column,N'_RecorderRRef' AS id_column
FROM sys.tables t WHERE t.is_ms_shipped=0
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_RecorderTRef' AND c.max_length=4)
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_RecorderRRef' AND c.max_length=16)
UNION ALL
SELECT t.name,N'_Recorder_RTRef',N'_Recorder_RRRef'
FROM sys.tables t WHERE t.is_ms_shipped=0
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_Recorder_RTRef' AND c.max_length=4)
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_Recorder_RRRef' AND c.max_length=16);
'@ 120
    foreach ($row in $recorderTables.Rows) {
        $table = [string]$row.table_name
        if (Is-DerivedTable $table) { continue }
        $typeColumn = Quote-SqlIdentifier ([string]$row.type_column)
        $idColumn = Quote-SqlIdentifier ([string]$row.id_column)
        $condition = "EXISTS (SELECT 1 FROM #TargetDocuments target WHERE src.$typeColumn=target.type_ref AND target.id=src.$idColumn)"
        Add-DeletePlanCondition $plans $table $condition 'recorder' 'recorder' 20
        $ownerTypeCode = "CONVERT(int,CONVERT(varbinary(4),src.$typeColumn))"
        Add-OwnerDescriptor $ownerDescriptors $table ("N'Document'+CONVERT(nvarchar(20),$ownerTypeCode)") $ownerTypeCode ("src.$idColumn") 'recorder'
    }

    $journalTables = Invoke-Table $connection $null @'
SELECT t.name AS table_name FROM sys.tables t WHERE t.is_ms_shipped=0
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_DocumentTRef' AND c.max_length=4)
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_DocumentRRef' AND c.max_length=16);
'@ 120
    foreach ($row in $journalTables.Rows) {
        $table = [string]$row.table_name
        $condition = 'EXISTS (SELECT 1 FROM #TargetDocuments target WHERE src.[_DocumentTRef]=target.type_ref AND target.id=src.[_DocumentRRef])'
        Add-DeletePlanCondition $plans $table $condition 'document-owner' 'document-journal' 10
        $ownerTypeCode = 'CONVERT(int,CONVERT(varbinary(4),src.[_DocumentTRef]))'
        Add-OwnerDescriptor $ownerDescriptors $table ("N'Document'+CONVERT(nvarchar(20),$ownerTypeCode)") $ownerTypeCode 'src.[_DocumentRRef]' 'document-journal'
    }

    $tabular = Invoke-Table $connection $null @'
SELECT t.name AS table_name,c.name AS owner_column,
       STUFF((SELECT N'|' + rc.name
              FROM sys.columns rc
              WHERE rc.object_id=t.object_id AND rc.max_length=16
                AND rc.name LIKE N'%RRef' AND rc.name<>c.name
                AND NOT EXISTS
                    (SELECT 1 FROM sys.columns tc
                     WHERE tc.object_id=t.object_id AND tc.max_length=4
                       AND tc.name LIKE N'%TRef'
                       AND LEFT(tc.name,LEN(tc.name)-4)+N'RRef'=rc.name)
              ORDER BY rc.column_id FOR XML PATH(N''),TYPE).value(N'.',N'nvarchar(max)'),1,1,N'') AS fixed_columns
FROM sys.tables t JOIN sys.columns c ON c.object_id=t.object_id
WHERE t.is_ms_shipped=0 AND t.name LIKE N'[_]%[_]VT%'
  AND c.name LIKE N'[_]%[_]IDRRef' AND c.max_length=16;
'@ 120
    foreach ($row in $tabular.Rows) {
        $table = [string]$row.table_name
        if ($table -notmatch '^_(?<token>[A-Za-z]+?)(?<code>[0-9]+)_VT[0-9]+(?:X[0-9]+)?$') { continue }
        $typeTable = "$($Matches.token)$($Matches.code)"
        $owner = Quote-SqlIdentifier ([string]$row.owner_column)
        $condition = "EXISTS (SELECT 1 FROM #TargetObjects target WHERE target.type_table=N'$typeTable' AND target.id=src.$owner)"
        Add-DeletePlanCondition $plans $table $condition 'tabular-owner' 'tabular' 20
        Add-OwnerDescriptor $ownerDescriptors $table ("N'$typeTable'") ([string][int]$Matches.code) ("src.$owner") 'tabular'
    }

    $objectTables = Invoke-Table $connection $null @'
SELECT t.name AS table_name FROM sys.tables t WHERE t.is_ms_shipped=0
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_IDRRef' AND c.max_length=16)
AND EXISTS(SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name=N'_Marked' AND c.max_length=1)
AND t.name NOT LIKE N'[_]DocumentJournal%';
'@ 120
    foreach ($row in $objectTables.Rows) {
        $table = [string]$row.table_name
        if ($table -notmatch '^_(?<token>[A-Za-z]+?)(?<code>[0-9]+)(?:X[0-9]+)?$') { continue }
        $typeTable = "$($Matches.token)$($Matches.code)"
        $condition = "EXISTS (SELECT 1 FROM #TargetObjects target WHERE target.type_table=N'$typeTable' AND target.id=src.[_IDRRef])"
        Add-DeletePlanCondition $plans $table $condition 'object-id' 'object' 30
        Add-OwnerDescriptor $ownerDescriptors $table ("N'$typeTable'") ([string][int]$Matches.code) 'src.[_IDRRef]' 'object'
    }

    # Fail before the first DELETE if a row that would survive under the
    # owner-only policy still points at an object in the delete manifest.
    # Such a target belongs to the KEEP/DELETE boundary and must be protected,
    # remapped or explicitly cleared; deleting only the referencing atomic row
    # is forbidden.
    $boundaryResults = [System.Collections.Generic.List[object]]::new()
    $simulatedCovered = [System.Collections.Generic.List[object]]::new()
    $boundaryIndex = 0
    foreach ($boundary in @($boundaryPlans.Values | Sort-Object table)) {
        $boundaryIndex++
        $table = [string]$boundary.table
        if (-not $plans.ContainsKey($table) -or $plans[$table].conditions.Count -eq 0) {
            throw "Atomic table '$table' has target references but no owner delete predicate. Classification is incomplete."
        }
        if (($boundaryIndex % 50) -eq 0 -or $boundaryIndex -eq $boundaryPlans.Count) {
            Write-Host ("KEEP boundary: table {0} of {1} ({2})" -f $boundaryIndex,$boundaryPlans.Count,$table)
        }
        $referencePredicate = '(' + (@($boundary.conditions) -join ' OR ') + ')'
        $ownerDeletePredicate = Get-DeletePlanPredicate $plans[$table] $protectedRegisterPredicates
        $countSql = "SELECT COUNT_BIG(*) FROM dbo.$(Quote-SqlIdentifier $table) AS src WHERE $referencePredicate AND NOT $ownerDeletePredicate;"
        $conflicts = [long](Invoke-Scalar $connection $null $countSql 1800)
        if ($conflicts -gt 0 -and $simulatedRuleRows.ContainsKey($table) -and $conflicts -le $simulatedRuleRows[$table]) {
            [void]$simulatedCovered.Add([pscustomobject]@{ table=$table; conflict_rows=$conflicts; rule_action_rows=$simulatedRuleRows[$table] })
            $conflicts = 0
        }
        if ($conflicts -gt 0) {
            [void]$boundaryResults.Add([pscustomobject]@{
                table=$table
                surviving_rows_referencing_delete_targets=$conflicts
                reference_paths=(@($boundary.reference_paths) | Sort-Object) -join '|'
                owner_delete_reasons=(@($plans[$table].reasons) | Sort-Object) -join '|'
            })
            if (-not $ownerDescriptors.ContainsKey($table)) {
                throw "Atomic table '$table' has no owner descriptor for KEEP-boundary evidence."
            }
            foreach ($reference in $boundary.references) {
              foreach ($ownerDescriptor in $ownerDescriptors[$table]) {
                $safePath = ([string]$reference.path).Replace("'","''")
                $safeOwnerRole = ([string]$ownerDescriptor.role).Replace("'","''")
                $evidenceSql = @"
INSERT #KeepBoundaryEvidence(type_table,type_code,id,source_table,reference_path,owner_type_table,owner_type_code,owner_id,owner_role)
SELECT DISTINCT target.type_table,target.type_code,target.id,N'$table',N'$safePath',
       $([string]$ownerDescriptor.type_table_expression),$([string]$ownerDescriptor.type_code_expression),$([string]$ownerDescriptor.id_expression),N'$safeOwnerRole'
FROM dbo.$(Quote-SqlIdentifier $table) AS src
JOIN #TargetObjects target ON $([string]$reference.target_join)
WHERE NOT $ownerDeletePredicate;
SELECT CONVERT(bigint,@@ROWCOUNT);
"@
                [void](Invoke-Scalar $connection $null $evidenceSql 1800)
              }
            }
        }
    }
    $boundaryResults | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'keep-boundary-conflicts.csv') -NoTypeInformation -Encoding UTF8
    $boundaryEvidence = Invoke-Table $connection $null @'
SELECT DISTINCT type_table,type_code,sys.fn_varbintohexstr(id) AS id_hex,source_table,reference_path,
       owner_type_table,owner_type_code,sys.fn_varbintohexstr(owner_id) AS owner_id_hex,owner_role
FROM #KeepBoundaryEvidence
ORDER BY type_table,id_hex,source_table,reference_path,owner_type_table,owner_id_hex;
'@ 300
    $boundaryEvidence.Rows | Select-Object type_table,type_code,id_hex,source_table,reference_path,owner_type_table,owner_type_code,owner_id_hex,owner_role |
        Export-Csv -LiteralPath (Join-Path $OutputDirectory 'keep-boundary-evidence.csv') -NoTypeInformation -Encoding UTF8
    $protectedTargets = Invoke-Table $connection $null @'
SELECT DISTINCT type_table,type_code,sys.fn_varbintohexstr(id) AS id_hex
FROM #KeepBoundaryEvidence
ORDER BY type_table,id_hex;
'@ 300
    $protectedTargets.Rows | Select-Object type_table,type_code,id_hex,@{Name='graph_depth';Expression={0}},@{Name='inclusion_reason';Expression={'keep-boundary'}} |
        Export-Csv -LiteralPath (Join-Path $OutputDirectory 'keep-protected-targets.csv') -NoTypeInformation -Encoding UTF8
    $boundarySummary = [ordered]@{
        generated_at_local=(Get-Date).ToString('s')
        server=$Server; database=$Database; read_only=$true
        atomic_tables=$atomicTables.Count; audited_tables=$boundaryPlans.Count
        conflict_tables=$boundaryResults.Count
        conflict_rows=[long](($boundaryResults | Measure-Object -Property surviving_rows_referencing_delete_targets -Sum).Sum)
        protected_targets=[long]$protectedTargets.Rows.Count
        protected_register_row_tables=$protectedRegisterLocations.Count
        protected_register_rows_sha256=$protectedRegisterRowsSha256
        preserved_organization_id_hex=$preservedOrganizationIdHex
        simulated_rule_action_tables=$simulatedCovered.Count
        simulated_rule_action_conflict_rows=[long](($simulatedCovered | Measure-Object -Property conflict_rows -Sum).Sum)
        result=if ($boundaryResults.Count -gt 0) { 'BLOCK' } elseif ($simulatedCovered.Count -gt 0) { 'PASS_SIMULATED' } else { 'PASS' }
    }
    $simulatedCovered | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'keep-boundary-simulated-rule-actions.csv') -NoTypeInformation -Encoding UTF8
    $boundarySummary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'keep-boundary-summary.json') -Encoding UTF8
    if ($BoundaryOnly) {
        $boundarySummary | ConvertTo-Json -Depth 5
        return
    }
    if ($boundaryResults.Count -gt 0) {
        throw "KEEP boundary gate blocked execution: $($boundaryResults.Count) atomic tables contain $($boundarySummary.conflict_rows) surviving rows that reference delete targets. See keep-boundary-conflicts.csv."
    }

    $actions = [System.Collections.Generic.List[object]]::new()
    $orderedPlans = @($plans.Values | Sort-Object phase,table)
    for ($planIndex=0; $planIndex -lt $orderedPlans.Count; $planIndex++) {
        $plan = $orderedPlans[$planIndex]
        if ($plan.conditions.Count -eq 0) { continue }
        $table = [string]$plan.table
        if (($planIndex % 50) -eq 0 -or $planIndex -eq ($orderedPlans.Count-1)) {
            Write-Host ("Delete plan: table {0} of {1} ({2})" -f ($planIndex+1),$orderedPlans.Count,$table)
        }
        $deleteFrom = "$(Quote-SqlIdentifier $table) AS src WHERE " + (Get-DeletePlanPredicate $plan $protectedRegisterPredicates)
        $tableBatchSize = Get-DeletePlanBatchSize $plan $BatchSize $RegisterBatchSize $JournalBatchSize $TabularBatchSize $ObjectBatchSize
        $actionName = "table:$table"
        $rows = Invoke-BatchedDelete -Connection $connection -Action $actionName -DeleteSql $deleteFrom -Log $log -BatchSize $tableBatchSize -Apply:$Apply
        [void]$actions.Add([pscustomobject]@{
            action=$actionName
            table=$table
            planned_or_deleted_rows=[long]$rows
            role=(@($plan.roles) | Sort-Object) -join '+'
            reasons=(@($plan.reasons) | Sort-Object) -join '+'
            batch_size=$tableBatchSize
            phase=$plan.phase
        })
    }

    $log | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'batch-log.csv') -NoTypeInformation -Encoding UTF8
    $result = [ordered]@{
        generated_at_local=(Get-Date).ToString('s'); read_only=(-not $Apply); dml_executed=[bool]$Apply
        server=$Server; database=$Database; batch_size=$BatchSize
        batch_sizes=[ordered]@{
            default=$BatchSize; register=$RegisterBatchSize; journal=$JournalBatchSize
            tabular=$TabularBatchSize; object=$ObjectBatchSize
        }
        execution_strategy='single-pass-per-physical-table'
        atomic_owner_only_policy=$true
        atomic_tables=$atomicTables.Count
        clear_reference_tabular_tables=$clearReferenceTabularTables.Count
        clear_reference_tabular_tables_manifest=if($ClearReferenceTabularTablesManifest){(Resolve-Path -LiteralPath $ClearReferenceTabularTablesManifest).Path}else{$null}
        clear_reference_tabular_tables_sha256=$clearReferenceTabularTablesSha256
        protected_register_row_tables=$protectedRegisterLocations.Count
        protected_register_rows_manifest=if($ProtectedRegisterRowsManifest){(Resolve-Path -LiteralPath $ProtectedRegisterRowsManifest).Path}else{$null}
        protected_register_rows_sha256=$protectedRegisterRowsSha256
        preserved_organization_id_hex=$preservedOrganizationIdHex
        keep_boundary_gate=$boundarySummary
        root_mode=$RootMode; object_manifest_sha256=$actualManifestSha256
        object_manifest=(Resolve-Path -LiteralPath $ObjectManifest).Path
        fixed_edges_manifest=(Resolve-Path -LiteralPath $FixedEdgesManifest).Path
        manifest_rows_loaded=$loaded; target_document_rows=$documentCount; target_organization_rows=$organizationCount
        preserved_organization_code=$preservedFilter.label; preserved_organization_in_manifest=$preservedInManifest
        derived_totals_touched=$false; info_register_policy='DELETE_ALL_LINKED_ROWS'
        actions=@($actions); batches=@($log)
        total_rows=[long](($actions | Measure-Object -Property planned_or_deleted_rows -Sum).Sum)
        duration_seconds=((Get-Date)-$startedOverall).TotalSeconds
    }
    $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
    $result | ConvertTo-Json -Depth 5
}
finally { $connection.Dispose() }
