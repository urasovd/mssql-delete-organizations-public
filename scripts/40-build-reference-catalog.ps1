[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$Database,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-reference-catalog'))
)
<#
    Read-only live catalog of every reference-shaped column in a 1C database.

    Earlier runs fed the executor a fixed-edge CSV produced by script 05, which
    only kept edges whose target type was present in the seed manifest at the
    time of the inventory.  Edges to types that entered the delete set later
    (nomenclature, deals, files) were silently missing, so both the boundary
    gate and the executor were blind to them (#4, #5).

    This script classifies every binary(16) *RRef column of every user table,
    independent of any manifest:
      id            _IDRRef of an object table
      tabular-owner _<Type><code>_IDRRef of a tabular part
      recorder      _RecorderRRef / _Recorder_RRRef
      journal-doc   _DocumentRRef of a document journal
      composite     *_RRRef with a matching *_RTRef type column
      fixed         single-type reference resolved through DBSchema
      owner/parent  _OwnerIDRRef / _ParentIDRRef resolved through DBSchema
      unresolved    nothing above matched; coverage gate fails

    Outputs:
      reference-catalog.csv      one row per physical column with its class
      fixed-reference-edges.csv  all fixed edges in the format of script 11
      type-names.csv             type token -> 1C logical name
      unresolved-columns.csv     columns without a class (must be empty)
      summary.json               counts and coverage_gate PASS|BLOCK
#>
$ErrorActionPreference = 'Stop'

function Invoke-Table([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds = 600) {
    $command = $Connection.CreateCommand()
    $command.CommandTimeout = $TimeoutSeconds
    $command.CommandText = "SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`n$Sql"
    try {
        $adapter = [System.Data.SqlClient.SqlDataAdapter]::new($command)
        try {
            $table = [System.Data.DataTable]::new()
            [void]$adapter.Fill($table)
            Write-Output -NoEnumerate $table
        }
        finally { $adapter.Dispose() }
    }
    finally { $command.Dispose() }
}

function Is-DerivedTable([string]$Name) {
    # Same list as script 11: platform-maintained slices, totals, aggregates.
    return ($Name -match '^_InfoRgSL' -or
            $Name -match '^_AccumRg(?:T|AT|Agg|St|Bf|Dl)' -or
            $Name -match '^_AccumRgOpt(?:$|X[0-9]+$)' -or
            $Name -match '^_AccRg(?:T|AT|CT)' -or
            $Name -match '^_AccRgOpt(?:$|X[0-9]+$)')
}

function Get-LogicalTable([string]$PhysicalTable) {
    # _Document175X1 -> Document175, _Document175_VT4415X1 -> Document175_VT4415
    $name = $PhysicalTable.TrimStart('_')
    if ($name -match '^(?<base>[A-Za-z][A-Za-z0-9_]*[0-9])X[0-9]+$') { $name = [string]$Matches.base }
    return $name
}

function Get-TableRole([string]$PhysicalTable) {
    $logical = Get-LogicalTable $PhysicalTable
    if (Is-DerivedTable $PhysicalTable) { return 'derived' }
    if ($logical -match '^[A-Za-z]+[0-9]+_VT[0-9]+$') { return 'tabular' }
    if ($logical -match '^(?:Acc|CKinds)[0-9]+_(?:ExtDim[0-9]+|LeadingCK|BaseCK|DisplacedCK)$') { return 'tabular' }
    if ($logical -match '^(?:STTModelsDesc|ConfigChngR|ExtsChngR)') { return 'system' }
    if ($logical -match '^DocumentJournal[0-9]+$') { return 'document-journal' }
    if ($logical -match '^(?:Document|Reference|Chrc|BPr|Task|Node|Acc|CKinds)[0-9]+$') { return 'object' }
    if ($logical -match '^InfoRg[0-9]+$') { return 'information-register' }
    if ($logical -match '^AccumRg[0-9]+$') { return 'accumulation-movement' }
    if ($logical -match '^AccRg(?:ED)?[0-9]+$') { return 'accounting-movement' }
    if ($logical -match '^CRg[0-9]+$') { return 'calculation-register' }
    if ($logical -match '^Const[0-9]+$') { return 'constant' }
    if ($logical -match '^Seq[0-9]+$') { return 'sequence' }
    if ($logical -match 'ChngR[0-9]*$') { return 'change-registration' }
    return 'other'
}

$assemblyPath = Join-Path $PSScriptRoot 'sqlclr\out\OneC.SqlClr.dll'
if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
    throw "Metadata parser is missing: $assemblyPath. Run sqlclr\Build-SqlClr.ps1 first."
}
if ($null -eq ('OneC.SqlClr.OneCMetadata' -as [type])) { Add-Type -LiteralPath $assemblyPath }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

$connection = [System.Data.SqlClient.SqlConnection]::new(
    "Server=$Server;Database=$Database;Integrated Security=True;Encrypt=False;Application Name=Codex-ReferenceCatalog")
try {
    $connection.Open()
    $identity = Invoke-Table $connection 'SELECT DB_NAME() AS database_name, d.state_desc FROM sys.databases d WHERE d.database_id=DB_ID();' 30
    if ([string]$identity.Rows[0].database_name -ne $Database) { throw "Connected to an unexpected database. Expected $Database." }
    if ([string]$identity.Rows[0].state_desc -ne 'ONLINE') { throw 'Database is not ONLINE.' }

    # DBNames: metadata id -> (token, code).  Type codes are unique across tokens.
    $dbNamesById = @{}
    $typeTableByCode = @{}
    $objectTokens = @('Document','Reference','Chrc','BPr','Task','Node','Acc','CKinds','Enum')
    # Types declared by configuration extensions (DBNames-Ext-*): their own
    # catalogs are not subject to COMPLEMENT patterns in 41, as in the 1C engine.
    $extensionTypes = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($row in (Invoke-Table $connection "SELECT FileName, BinaryData FROM dbo.Params WHERE FileName=N'DBNames' OR FileName LIKE N'DBNames-Ext-%';" 300).Rows) {
        $fromExtension = [string]$row.FileName -ne 'DBNames'
        foreach ($item in [OneC.SqlClr.OneCMetadata]::ParseDbNames([System.Data.SqlTypes.SqlBytes]::new([byte[]]$row.BinaryData))) {
            $id = $item.MetadataId.ToString().ToLowerInvariant()
            if (-not $dbNamesById.ContainsKey($id)) { $dbNamesById[$id] = [System.Collections.Generic.List[object]]::new() }
            $dbNamesById[$id].Add([pscustomobject]@{ token=[string]$item.Token; code=[int]$item.Code })
            if ($objectTokens -contains [string]$item.Token) {
                $typeTableByCode[[int]$item.Code] = [string]$item.Token + [int]$item.Code
                if ($fromExtension) { [void]$extensionTypes.Add([string]$item.Token + [int]$item.Code) }
            }
        }
    }

    $typeTableById = @{}
    foreach ($id in $dbNamesById.Keys) {
        foreach ($dbName in $dbNamesById[$id]) {
            if ($objectTokens -contains $dbName.token) { $typeTableById[$id] = [string]$dbName.token + [int]$dbName.code }
        }
    }

    # Config: logical names of types and fields, for readable reports, and the
    # owner lists of catalogs (designer tab "Владельцы").  An owner stored as
    # _OwnerID_RTRef/_OwnerID_RRRef is "any reference" in the storage schema,
    # so its real target types come only from this list.
    $typeNames = @{}
    $fieldNames = @{}
    $ownersByType = @{}
    $configCommand = $connection.CreateCommand()
    $configCommand.CommandTimeout = 900
    $configCommand.CommandText = @'
SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
SELECT FileName, BinaryData, 0 AS from_extension FROM dbo.Config
WHERE LEN(FileName)=36 AND SUBSTRING(FileName,9,1)=N'-' AND SUBSTRING(FileName,14,1)=N'-'
  AND SUBSTRING(FileName,19,1)=N'-' AND SUBSTRING(FileName,24,1)=N'-';
-- Objects of configuration extensions (their own catalogs, documents,
-- tabular parts and fields) are described only in ConfigCAS, stored by
-- content hash.  Without them policies by logical name cannot reach them
-- (found on an extension's own catalog).
IF OBJECT_ID(N'dbo.ConfigCAS') IS NOT NULL
    SELECT FileName, BinaryData, 1 AS from_extension FROM dbo.ConfigCAS;
'@
    $reader = $configCommand.ExecuteReader()
    try {
        # Config first: its names win over the extensions' (ContainsKey below).
        do { while ($reader.Read()) {
            $fileId = $reader.GetString(0).ToLowerInvariant()
            $fileBytes = [System.Data.SqlTypes.SqlBytes]::new([byte[]]$reader.GetValue(1))
            try { $names = @([OneC.SqlClr.OneCMetadata]::ParseConfigNames($fileBytes)) }
            catch { continue }
            if ([string]$typeTableById[$fileId] -match '^Reference[0-9]+$') {
                $owners = [OneC.SqlClr.OneCMetadata]::ParseCatalogOwners($fileBytes, $fileId)
                if ($null -ne $owners) {
                    $ownersByType[$typeTableById[$fileId]] = @($owners | ForEach-Object {
                        if ($typeTableById.ContainsKey($_)) { $typeTableById[$_] } else { "Unknown:$_" } })
                }
            }
            foreach ($nameRow in $names) {
                $id = $nameRow.MetadataId.ToString().ToLowerInvariant()
                if (-not $dbNamesById.ContainsKey($id)) { continue }
                foreach ($dbName in $dbNamesById[$id]) {
                    $key = [string]$dbName.token + [int]$dbName.code
                    if ($dbName.token -eq 'Fld') { if (-not $fieldNames.ContainsKey($key)) { $fieldNames[$key] = [string]$nameRow.LogicalName } }
                    elseif (-not $typeNames.ContainsKey($key)) { $typeNames[$key] = [string]$nameRow.LogicalName }
                }
            }
        } } while ($reader.NextResult())
    }
    finally { $reader.Dispose(); $configCommand.Dispose() }

    # Storage schema: SchemaStorage holds the main schema (SchemaID=0) and one
    # row per configuration extension; DBSchema is the main schema only and
    # misses every extension field (_Document175X1._Fld71751RRef and others).
    $schemaRows = Invoke-Table $connection @'
IF OBJECT_ID(N'dbo.SchemaStorage') IS NOT NULL
    SELECT SchemaID AS schema_id, CurrentSchema AS schema_bytes FROM dbo.SchemaStorage;
ELSE
    SELECT 0 AS schema_id, SerializedData AS schema_bytes FROM dbo.DBSchema;
'@ 300
    # logical table|physical column -> reference targets of that field
    $fieldTargets = @{}
    foreach ($schemaRow in $schemaRows.Rows) {
        foreach ($item in [OneC.SqlClr.OneCMetadata]::ParseReferenceFieldTargets(
                [System.Data.SqlTypes.SqlBytes]::new([byte[]]$schemaRow.schema_bytes))) {
            $field = [string]$item.SourceField
            if ($field -eq 'ID') { continue }
            $column = if ([int]$item.TypeCount -eq 1) { "_${field}RRef" } else { "_${field}_RRRef" }
            $key = [string]$item.SourceTable + '|' + $column
            if (-not $fieldTargets.ContainsKey($key)) {
                $fieldTargets[$key] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            }
            [void]$fieldTargets[$key].Add([string]$item.TargetTable)
        }
    }

    $columns = Invoke-Table $connection @'
SELECT t.name AS table_name, c.name AS column_name,
       CASE WHEN EXISTS(SELECT 1 FROM sys.columns m WHERE m.object_id=t.object_id AND m.name=N'_Marked' AND m.max_length=1) THEN 1 ELSE 0 END AS has_marked,
       (SELECT TOP (1) tc.name FROM sys.columns tc
         WHERE tc.object_id=t.object_id AND tc.max_length=4
           AND tc.name IN (LEFT(c.name,LEN(c.name)-5)+N'_RTRef', LEFT(c.name,LEN(c.name)-4)+N'TRef')) AS type_column
FROM sys.tables t
JOIN sys.schemas s ON s.schema_id=t.schema_id AND s.name=N'dbo'
JOIN sys.columns c ON c.object_id=t.object_id
WHERE t.is_ms_shipped=0 AND t.name LIKE N'[_]%' AND c.max_length=16 AND c.name LIKE N'%RRef'
ORDER BY t.name, c.column_id;
'@ 300

    $catalog = [System.Collections.Generic.List[object]]::new()
    $fixedEdges = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $columns.Rows) {
        $table = [string]$row.table_name
        $column = [string]$row.column_name
        if ($table -notmatch '^_[A-Za-z0-9_]+$' -or $column -notmatch '^_[A-Za-z0-9_]+$') { throw "Unsafe physical name: $table.$column" }
        $logical = Get-LogicalTable $table
        $role = Get-TableRole $table
        $typeColumn = if ($row.IsNull('type_column')) { $null } else { [string]$row.type_column }
        $kind = 'unresolved'
        $target = $null
        $targets = $fieldTargets["$logical|$column"]
        if (-not $targets) { $targets = $fieldTargets["$($table.TrimStart('_'))|$column"] }
        if (-not $targets -and $column -match '^_(?<f>[A-Za-z0-9]+)RRef$') {
            # Composite system fields keep the short spelling: the schema field
            # Recorder with several types is stored as _RecorderTRef/_RecorderRRef.
            $targets = $fieldTargets["$logical|_$($Matches.f)_RRRef"]
        }
        if ($column -in @('_OwnerIDRRef','_OwnerID_RRRef') -and $logical -match '^Reference[0-9]+$') {
            # Owner targets come from the catalog's owner list.  Without a list
            # (catalogs of configuration extensions are not in Config) a fixed
            # single-type owner keeps its unambiguous schema target, while a
            # composite owner is left unresolved.
            if ($ownersByType.ContainsKey($logical) -and $ownersByType[$logical].Count -gt 0) {
                $targets = [System.Collections.Generic.HashSet[string]]::new([string[]]$ownersByType[$logical], [System.StringComparer]::OrdinalIgnoreCase)
            }
            elseif ($column -eq '_OwnerID_RRRef') { $targets = $null }
        }
        $namedTargets = @(if ($targets) { $targets | Where-Object { $_ } })

        if ($column -eq '_IDRRef') { $kind = 'id' }
        elseif ($role -eq 'system') { $kind = 'system' }
        elseif ($role -eq 'tabular' -and $column -match '^_[A-Za-z]+[0-9]+_IDRRef$') { $kind = 'tabular-owner' }
        elseif ($column -in @('_RecorderRRef','_Recorder_RRRef')) { $kind = 'recorder' }
        elseif ($role -eq 'document-journal' -and $column -eq '_DocumentRRef') { $kind = 'journal-doc' }
        elseif ($typeColumn) { $kind = 'composite' }
        elseif ($namedTargets.Count -eq 1 -and -not $targets.Contains('')) {
            # A plain fixed reference, or a composite field whose only
            # reference type is single: 1C then omits the _RTRef column.
            $target = $namedTargets[0]
            $kind = switch -Regex ($column) { '^_OwnerIDRRef$' { 'owner' } '^_ParentIDRRef$' { 'parent' } '_RRRef$' { 'composite-single' } default { 'fixed' } }
        }
        elseif ($namedTargets.Count -gt 1) { $kind = 'ambiguous' }
        elseif ($column -eq '_ParentIDRRef' -and $logical -match '^(?:Reference|Chrc|Acc|CKinds|BPr|Task)[0-9]+$') {
            $kind = 'parent'; $target = $logical
        }
        if ($column -in @('_OwnerIDRRef','_OwnerID_RRRef') -and $logical -match '^Reference[0-9]+$' -and -not $targets) { $kind = 'unresolved' }
        if ($kind -in @('recorder','journal-doc') -and -not $typeColumn) {
            # A register with a single recorder type (or a journal of one
            # document type) has no type column; script 11 recognises recorders
            # only by the type/id pair, so such columns must reach it as edges.
            if ($namedTargets.Count -eq 1) { $target = $namedTargets[0] } else { $kind = 'unresolved' }
        }

        $fieldCode = if ($column -match '^_(?<f>Fld[0-9]+)') { [string]$Matches.f } else { $null }
        $catalog.Add([pscustomobject]@{
            physical_table=$table; column=$column; kind=$kind; target_table=$target; type_column=$typeColumn
            # composite targets from the schema; '*' means any reference type
            targets=if ($targets) { (@($targets | ForEach-Object { if ($_) { $_ } else { '*' } } | Sort-Object) -join ';') } else { $null }
            table_role=$role; logical_table=$logical
            table_name=if ($logical -match '^(?<b>[A-Za-z]+[0-9]+)') { $typeNames[[string]$Matches.b] } else { $null }
            field_name=if ($fieldCode) { $fieldNames[$fieldCode] } else { $null }
        })
        if ($target -and $kind -in @('fixed','composite-single','owner','parent','recorder','journal-doc') -and $role -ne 'derived') {
            $fixedEdges.Add([pscustomobject]@{
                source_table=$logical; source_field=($column.TrimStart('_') -replace 'RRef$',''); target_table=$target
                physical_table=$table; physical_column=$column; source_has_marked_column=[int]$row.has_marked
            })
        }
    }

    $catalogPath = Join-Path $OutputDirectory 'reference-catalog.csv'
    $edgesPath = Join-Path $OutputDirectory 'fixed-reference-edges.csv'
    $unresolvedPath = Join-Path $OutputDirectory 'unresolved-columns.csv'
    $typeNamesPath = Join-Path $OutputDirectory 'type-names.csv'
    $catalog | Export-Csv -LiteralPath $catalogPath -NoTypeInformation -Encoding UTF8
    $fixedEdges | Sort-Object physical_table,physical_column,target_table |
        Export-Csv -LiteralPath $edgesPath -NoTypeInformation -Encoding UTF8
    $unresolved = @($catalog | Where-Object { $_.kind -in @('unresolved','ambiguous') -and $_.table_role -ne 'derived' })
    $unresolved | Export-Csv -LiteralPath $unresolvedPath -NoTypeInformation -Encoding UTF8
    @($typeNames.GetEnumerator() | Sort-Object Name | ForEach-Object { [pscustomobject]@{ type_table=$_.Name; logical_name=$_.Value; extension=[int]$extensionTypes.Contains($_.Name) } }) |
        Export-Csv -LiteralPath $typeNamesPath -NoTypeInformation -Encoding UTF8
    # Every field, not only references: policies address non-reference
    # fields too (Пользователи.ИдентификаторПользователяИБ).
    @($fieldNames.GetEnumerator() | Sort-Object Name | ForEach-Object { [pscustomobject]@{ field=$_.Name; logical_name=$_.Value } }) |
        Export-Csv -LiteralPath (Join-Path $OutputDirectory 'field-names.csv') -NoTypeInformation -Encoding UTF8

    $summary = [ordered]@{
        generated_at_local=(Get-Date).ToString('s'); server=$Server; database=$Database; read_only=$true
        reference_columns=$catalog.Count
        by_kind=[ordered]@{}
        fixed_edges=$fixedEdges.Count
        unresolved_non_derived_columns=$unresolved.Count
        coverage_gate=if ($unresolved.Count -eq 0) { 'PASS' } else { 'BLOCK' }
        catalog=$catalogPath
        fixed_edges_manifest=$edgesPath
        fixed_edges_sha256=(Get-FileHash -LiteralPath $edgesPath -Algorithm SHA256).Hash.ToLowerInvariant()
        unresolved_columns=$unresolvedPath
    }
    foreach ($group in @($catalog | Group-Object kind | Sort-Object Name)) { $summary.by_kind[$group.Name] = $group.Count }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
    $summary | ConvertTo-Json -Depth 5
}
finally { $connection.Dispose() }
