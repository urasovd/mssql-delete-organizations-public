[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$Database,
    # Kept organizations by reference (see lib/Organizations.ps1), one or more.
    [Alias('OrganizationId')] [string[]]$Organization,
    # Or by code, where the organization catalog has one.
    [string[]]$OrganizationCode,
    [Parameter(Mandatory)] [string]$CatalogDirectory,
    [string]$OrganizationCatalogName = 'Организации',
    [ValidateRange(1,200)] [int]$MaxRounds = 60,
    [ValidateRange(1,500)] [int]$StatementsPerBatch = 100,
    # Catalog profile: policies\catalogs-<profile>.csv, common to all configurations.
    [ValidateSet('base','aggressive')] [string]$CatalogProfile = 'base',
    # Policy layers: the common file first, then run layers.  A row of a later
    # layer replaces the rows of earlier layers for the same object (catalogs)
    # or object + table part + field (boundary rules); role/action NONE drops it.
    [string[]]$CatalogPolicyPath = @(Join-Path $PSScriptRoot "..\policies\catalogs-$CatalogProfile.csv"),
    [string[]]$BoundaryRulesPath = @(Join-Path $PSScriptRoot '..\policies\boundary-rules.csv'),
    [string[]]$CascadeProtectedTypeNames = @('Пользователи','ВнешниеПользователи','ГруппыПользователей','ГруппыВнешнихПользователей','ПрофилиГруппДоступа'),
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-keep-closure'))
)
<#
    Read-only KEEP closure for the kept organizations, and DELETE as its
    complement.  "Kept organization" below means any of them.

    Earlier iterations started from the organizations being deleted and walked
    forward.  Everything that was not reachable from a foreign organization
    was missed: documents without an organization (#2), deals and clients (#4).
    Here the direction is reversed:

      1. Object inventory.  Every object row (header and X-extension tables) is
         classified by its organization-typed fields:
           has_keep     some organization field equals the kept organization
           has_foreign  some organization field holds another organization
         An organization field is any column whose schema type is the
         organization catalog (fixed, single-type composite, owner, or a
         composite listing that catalog), not only fields named "Организация".

      2. KEEP seeds: the kept organization; objects with has_keep; recorders
         of register rows of the kept organization; predefined items; and the
         references held by constants and by register rows of the organization.

      3. KEEP closure.  Rows that survive with a KEEP object (header, X tables,
         tabular parts, movements it recorded, journal rows) keep every object
         they reference.  A referenced foreign organization or foreign-only
         object is not pulled in: it is written to boundary-references.csv and
         needs an explicit CLEAR/REMAP/KEEP decision.

      4. DELETE = existing objects not in KEEP that are documents, other
         organizations, or objects owned only by other organizations
         (predefined excluded).  Catalogs without an organization stay: they
         belong to the reference-data stage (#1).

      5. Cascade.  An object outside KEEP and DELETE whose header or tabular
         part references a DELETE object is not needed by the kept
         organization (it would be in KEEP otherwise) and would keep a broken
         reference: attached files of deleted documents, access keys, deal
         satellites.  Such objects are written to a separate cascade manifest
         with the field that triggered them; delete-manifest-with-cascade.csv
         is the union.  Types listed in -CascadeProtectedTypeNames (users,
         access profiles) are never cascaded and are reported as blocked.

    Requires a reference catalog from script 40 with coverage_gate=PASS.
    The database is never modified; only #temp tables are created.
#>
$ErrorActionPreference = 'Stop'
$zeroId = '0x00000000000000000000000000000000'
. (Join-Path $PSScriptRoot 'lib\Organizations.ps1')

function Quote-SqlIdentifier([string]$Name) {
    if ($Name -notmatch '^_[A-Za-z0-9_]+$') { throw "Unsafe SQL identifier: $Name" }
    return '[' + $Name + ']'
}

function Invoke-NonQuery([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds = 3600) {
    $command = $Connection.CreateCommand()
    $command.CommandTimeout = $TimeoutSeconds
    $command.CommandText = "SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`n$Sql"
    try { [void]$command.ExecuteNonQuery() } finally { $command.Dispose() }
}

function Invoke-Scalar([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds = 3600) {
    $command = $Connection.CreateCommand()
    $command.CommandTimeout = $TimeoutSeconds
    $command.CommandText = "SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`n$Sql"
    try { return $command.ExecuteScalar() } finally { $command.Dispose() }
}

function Invoke-Table([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds = 3600) {
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

function Invoke-Batched([System.Data.SqlClient.SqlConnection]$Connection, [System.Collections.Generic.List[string]]$Statements, [string]$Label) {
    for ($offset = 0; $offset -lt $Statements.Count; $offset += $StatementsPerBatch) {
        $count = [Math]::Min($StatementsPerBatch, $Statements.Count - $offset)
        Invoke-NonQuery $Connection ($Statements.GetRange($offset, $count) -join "`n")
        Write-Host ("{0}: {1} of {2} statements" -f $Label, ($offset + $count), $Statements.Count)
    }
}

function Export-QueryCsv([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [string]$Path) {
    # Streams large result sets; every value is quoted like Export-Csv so the
    # manifest header and rows stay byte-compatible with script 11.
    $command = $Connection.CreateCommand()
    $command.CommandTimeout = 3600
    $command.CommandText = "SET NOCOUNT ON; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`n$Sql"
    $writer = [System.IO.StreamWriter]::new($Path, $false, [System.Text.UTF8Encoding]::new($false))
    $rows = [int64]0
    try {
        $reader = $command.ExecuteReader()
        try {
            $names = for ($i = 0; $i -lt $reader.FieldCount; $i++) { '"' + $reader.GetName($i) + '"' }
            $writer.WriteLine($names -join ',')
            $values = [string[]]::new($reader.FieldCount)
            while ($reader.Read()) {
                for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                    $values[$i] = if ($reader.IsDBNull($i)) { '""' } else { '"' + ([string]$reader.GetValue($i)).Replace('"','""') + '"' }
                }
                $writer.WriteLine($values -join ',')
                $rows++
            }
        }
        finally { $reader.Dispose() }
    }
    finally { $writer.Dispose(); $command.Dispose() }
    return $rows
}

function Get-TypeCode([string]$TypeTable) {
    if ($TypeTable -notmatch '^[A-Za-z]+(?<code>[0-9]+)$') { throw "Invalid type token: $TypeTable" }
    return [int]$Matches.code
}

# ---------------------------------------------------------------- catalog
$catalogSummaryPath = Join-Path $CatalogDirectory 'summary.json'
$catalogSummary = Get-Content -LiteralPath $catalogSummaryPath -Raw | ConvertFrom-Json
if ($catalogSummary.coverage_gate -ne 'PASS') { throw "Reference catalog coverage gate is $($catalogSummary.coverage_gate); rebuild it with script 40." }
if ($catalogSummary.database -ne $Database) { throw "Reference catalog was built for $($catalogSummary.database), not $Database." }
$catalog = @(Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'reference-catalog.csv') | Where-Object { $_.table_role -ne 'derived' })
$typeNames = @{}
foreach ($row in Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'type-names.csv')) { $typeNames[$row.type_table] = $row.logical_name }

$orgTypes = @($typeNames.GetEnumerator() | Where-Object { $_.Value -eq $OrganizationCatalogName -and $_.Name -match '^Reference[0-9]+$' })
if ($orgTypes.Count -ne 1) { throw "Expected one catalog named '$OrganizationCatalogName'; found $($orgTypes.Count)." }
$orgType = [string]$orgTypes[0].Name
$orgCode = Get-TypeCode $orgType
$orgTable = '_' + $orgType
$organizationFilter = Get-OrganizationFilter $Organization $OrganizationCode
# Organization fields are tested against #KeepOrg, filled once the
# organizations are looked up.
$isKeptSql = { param($Column) "EXISTS(SELECT 1 FROM #KeepOrg ko WHERE ko.id=$Column)" }

# Object tables per type (base + X extensions).
$objectTablesByType = @{}
foreach ($row in $catalog | Where-Object { $_.kind -eq 'id' -and $_.table_role -eq 'object' }) {
    if (-not $objectTablesByType.ContainsKey($row.logical_table)) { $objectTablesByType[$row.logical_table] = [System.Collections.Generic.List[string]]::new() }
    $objectTablesByType[$row.logical_table].Add($row.physical_table)
}
$objectTypeByCode = @{}
foreach ($type in $objectTablesByType.Keys) { $objectTypeByCode[(Get-TypeCode $type)] = $type }

# Reference columns per physical table, as VALUES rows (type_code, id, column).
$refKinds = @('fixed','composite-single','owner','parent','composite')
$refValuesByTable = @{}
$orgPredicatesByTable = @{}
foreach ($row in $catalog | Where-Object { $_.kind -in $refKinds }) {
    $table = [string]$row.physical_table
    $column = Quote-SqlIdentifier ([string]$row.column)
    $value = $null
    $isOrgField = $false
    if ($row.kind -eq 'composite') {
        $typeColumn = Quote-SqlIdentifier ([string]$row.type_column)
        $value = "(CONVERT(int,CONVERT(varbinary(4),src.$typeColumn)),src.$column,N'$($row.column)')"
        # Owner targets come from the catalog owner list (script 40); fields
        # declared as "any reference" ('*') are not an organization of the row.
        $isOrgField = (@(([string]$row.targets) -split ';') -contains $orgType)
        if ($isOrgField) {
            $orgPredicate = [pscustomobject]@{
                keep="(src.$typeColumn=CONVERT(binary(4),$orgCode) AND $(& $isKeptSql "src.$column"))"
                foreign="(src.$typeColumn=CONVERT(binary(4),$orgCode) AND src.$column<>$zeroId AND NOT $(& $isKeptSql "src.$column"))"
                column=[string]$row.column; type_column=[string]$row.type_column
            }
        }
    }
    else {
        if (-not $objectTypeByCode.ContainsKey((Get-TypeCode ([string]$row.target_table)))) { continue }
        $value = "($(Get-TypeCode ([string]$row.target_table)),src.$column,N'$($row.column)')"
        $isOrgField = ([string]$row.target_table -eq $orgType)
        if ($isOrgField) {
            $orgPredicate = [pscustomobject]@{
                keep="($(& $isKeptSql "src.$column"))"
                foreign="(src.$column<>$zeroId AND NOT $(& $isKeptSql "src.$column"))"
                column=[string]$row.column; type_column=$null
            }
        }
    }
    if (-not $refValuesByTable.ContainsKey($table)) { $refValuesByTable[$table] = [System.Collections.Generic.List[string]]::new() }
    $refValuesByTable[$table].Add($value)
    if ($isOrgField) {
        if (-not $orgPredicatesByTable.ContainsKey($table)) { $orgPredicatesByTable[$table] = [System.Collections.Generic.List[object]]::new() }
        $orgPredicatesByTable[$table].Add($orgPredicate)
    }
}

# Rows that live and die with an object: tabular parts, movements, journals.
$dependentSourcesByType = @{}
function Add-DependentSource([string]$Type, [string]$Table, [string]$KeyColumn, [string]$TypeColumn) {
    if (-not $dependentSourcesByType.ContainsKey($Type)) { $dependentSourcesByType[$Type] = [System.Collections.Generic.List[object]]::new() }
    $dependentSourcesByType[$Type].Add([pscustomobject]@{ table=$Table; key=$KeyColumn; type_column=$TypeColumn })
}
foreach ($row in $catalog | Where-Object { $_.kind -eq 'tabular-owner' }) {
    if ($row.logical_table -notmatch '^(?<owner>[A-Za-z]+[0-9]+)_') { continue }
    Add-DependentSource ([string]$Matches.owner) $row.physical_table $row.column $null
}
foreach ($row in $catalog | Where-Object { $_.kind -in @('recorder','journal-doc') -and $_.table_role -notin @('change-registration','sequence') }) {
    $owners = if ($row.type_column) { @(([string]$row.targets) -split ';' | Where-Object { $_ -and $_ -ne '*' }) } else { @([string]$row.target_table) }
    foreach ($owner in $owners) { Add-DependentSource $owner $row.physical_table $row.column $row.type_column }
}

$roleByTable = @{}
$recorderByTable = @{}
foreach ($row in $catalog) {
    $roleByTable[[string]$row.physical_table] = [string]$row.table_role
    if ($row.kind -eq 'recorder') { $recorderByTable[[string]$row.physical_table] = $row }
}
$registerRoles = @('information-register','accumulation-movement','accounting-movement','calculation-register')
$registerOrgTables = @($orgPredicatesByTable.Keys | Where-Object { $roleByTable[$_] -in $registerRoles } | Sort-Object)
$constantTables = @($refValuesByTable.Keys | Where-Object { $_ -match '^_Const[0-9]+(?:X[0-9]+)?$' } | Sort-Object)


$fieldNameByCode = @{}
$fieldNamesPath = Join-Path $CatalogDirectory 'field-names.csv'
if (Test-Path -LiteralPath $fieldNamesPath -PathType Leaf) {
    foreach ($row in Import-Csv -LiteralPath $fieldNamesPath) { $fieldNameByCode[$row.field] = $row.logical_name }
}

# Logical object names ("Справочник.X", "ПланОбмена.*", "Справочник.Классификатор*")
# resolve to type tokens.  A trailing/leading * is a name pattern.
$tokenByKind = @{ 'Справочник'='Reference'; 'Документ'='Document'; 'ПланВидовХарактеристик'='Chrc'; 'ПланОбмена'='Node'
                  'БизнесПроцесс'='BPr'; 'Задача'='Task'; 'РегистрСведений'='InfoRg'; 'РегистрНакопления'='AccumRg'; 'РегистрБухгалтерии'='AccRg'
                  'ПланВидовРасчета'='CKinds'; 'ГраницыПоследовательности'='SeqB' }
function Resolve-ObjectPattern([string]$Object, [switch]$AllowMissing) {
    if ($Object -notmatch '^(?<kind>[^.]+)\.(?<name>.+)$' -or -not $tokenByKind.ContainsKey($Matches.kind)) { throw "Unsupported object name: $Object" }
    $token = $tokenByKind[$Matches.kind]
    $name = [string]$Matches.name
    $types = @($typeNames.GetEnumerator() | Where-Object { $_.Name -match ('^' + $token + '[0-9]+$') -and $_.Value -like $name } | ForEach-Object { [string]$_.Name })
    if ($name -notmatch '\*' -and $types.Count -ne 1) {
        # One policy serves several configurations: an object this
        # configuration does not have is skipped and listed in the summary.
        if ($AllowMissing -and $types.Count -eq 0) { Write-Warning "Policy object not found, skipped: $Object"; [void]$skippedPolicyObjects.Add($Object); return @() }
        throw "Expected one object '$Object', found $($types.Count)."
    }
    return $types
}

# Catalog policy: COMPLEMENT types lose everything outside KEEP; KEEP_ROOT
# types (optionally only rows meeting a condition) seed KEEP; PROTECT types are
# neither deleted nor cascaded.  KEEP_ROOT and PROTECT win over COMPLEMENT.
$complementTypes = [System.Collections.Generic.HashSet[string]]::new()
$protectTypes = [System.Collections.Generic.HashSet[string]]::new()
$keepRoots = [System.Collections.Generic.List[object]]::new()
$keepWith = [System.Collections.Generic.List[object]]::new()
function Import-PolicyLayers([string[]]$Paths, [string[]]$KeyColumns, [string]$RoleColumn) {
    $rows = [System.Collections.Generic.List[object]]::new()
    $files = [System.Collections.Generic.List[object]]::new()
    foreach ($path in @($Paths | Where-Object { $_ })) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Policy layer not found: $path" }
        $layer = @(Import-Csv -LiteralPath $path -Delimiter ';')
        $keyOf = { param($Row) (@(foreach ($column in $KeyColumns) { [string]$Row.$column }) -join '|') }
        $keys = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($row in $layer) { [void]$keys.Add((& $keyOf $row)) }
        $kept = @(foreach ($row in $rows) { if (-not $keys.Contains((& $keyOf $row))) { $row } })
        $rows.Clear(); foreach ($row in $kept) { $rows.Add($row) }
        foreach ($row in $layer) { if ([string]$row.$RoleColumn -ne 'NONE') { $rows.Add($row) } }
        $files.Add([ordered]@{ path=(Resolve-Path -LiteralPath $path).Path; sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant(); rows=$layer.Count })
    }
    return [pscustomobject]@{ rows=@($rows); files=@($files) }
}
$skippedPolicyObjects = [System.Collections.Generic.SortedSet[string]]::new()
$catalogLayers = Import-PolicyLayers $CatalogPolicyPath @('object') 'role'
$catalogPolicyRows = @($catalogLayers.rows)
if ($catalogPolicyRows.Count) {
    foreach ($policy in $catalogPolicyRows) {
        $types = @(Resolve-ObjectPattern ([string]$policy.object) -AllowMissing)
        switch ([string]$policy.role) {
            'COMPLEMENT' { foreach ($type in $types) { [void]$complementTypes.Add($type) } }
            'PROTECT' { foreach ($type in $types) { [void]$protectTypes.Add($type) } }
            'KEEP_ROOT' { foreach ($type in $types) { $keepRoots.Add([pscustomobject]@{ type=$type; condition=[string]$policy.condition }) } }
            'KEEP_WITH' {
                if (-not $policy.condition) { throw "KEEP_WITH needs a field in condition: $($policy.object)." }
                foreach ($type in $types) { $keepWith.Add([pscustomobject]@{ type=$type; field=[string]$policy.condition }) }
            }
            default { throw "Unknown catalog policy role '$($policy.role)' for $($policy.object)." }
        }
    }
}
# Own catalogs of configuration extensions are not complemented (decision
# 08.11.2026, the same in the 1C engine): their data goes only with the
# extension.  40 marks them in type-names.csv (column extension).
$extensionComplementSkipped = [System.Collections.Generic.SortedSet[string]]::new()
foreach ($row in Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'type-names.csv')) {
    if ([string]$row.extension -eq '1' -and $complementTypes.Remove([string]$row.type_table)) { [void]$extensionComplementSkipped.Add([string]$row.logical_name) }
}
# KEEP_WITH: an object or register row stays with the object its field points
# to - attached files with ВладелецФайла, ХранилищеФайлов rows with Файл.  The
# reference goes from the row to KEEP, so forward closure never reaches it; the
# row is added as a dependent source of the referenced type, and an object row
# also joins KEEP itself (self_code).
foreach ($spec in $keepWith) {
    $tables = if ($objectTablesByType.ContainsKey($spec.type)) { @($objectTablesByType[$spec.type]) } else { @($catalog | Where-Object { $_.logical_table -eq $spec.type } | ForEach-Object physical_table | Select-Object -Unique) }
    $selfCode = if ($objectTablesByType.ContainsKey($spec.type)) { Get-TypeCode $spec.type } else { $null }
    $columns = @($catalog | Where-Object { $_.physical_table -in $tables -and $_.field_name -eq $spec.field -and $_.kind -in @('fixed','composite-single','composite') })
    if ($columns.Count -eq 0) { throw "KEEP_WITH field '$($spec.field)' not found in $($typeNames[$spec.type])." }
    foreach ($column in $columns) {
        $owners = if ($column.kind -eq 'composite') { @(([string]$column.targets) -split ';' | Where-Object { $_ }) } else { @([string]$column.target_table) }
        if ($owners -contains '*') { $owners = @($objectTablesByType.Keys) }
        $typeColumn = if ($column.kind -eq 'composite') { [string]$column.type_column } else { $null }
        foreach ($owner in $owners) {
            Add-DependentSource $owner $column.physical_table $column.column $typeColumn
            $dependentSourcesByType[$owner][$dependentSourcesByType[$owner].Count - 1] | Add-Member -NotePropertyName self_code -NotePropertyValue $selfCode
        }
    }
}
# An unconditional KEEP_ROOT keeps the whole type; a conditional one keeps only
# matching rows, the rest of the type stays subject to COMPLEMENT.
foreach ($root in $keepRoots) { if (-not $root.condition) { [void]$protectTypes.Add($root.type) } }
[void]$protectTypes.Add($orgType)
$complementCodes = @($complementTypes | Where-Object { -not $protectTypes.Contains($_) -and $objectTablesByType.ContainsKey($_) } | ForEach-Object { Get-TypeCode $_ })
$protectCodes = @($protectTypes | Where-Object { $objectTablesByType.ContainsKey($_) } | ForEach-Object { Get-TypeCode $_ })
$complementList = if ($complementCodes.Count) { $complementCodes -join ',' } else { '-1' }

# Boundary rules: a matched reference is resolved by the rule (cleared by
# script 42) instead of holding its target in KEEP or blocking the run.
$boundaryLayers = Import-PolicyLayers $BoundaryRulesPath @('object','table_part','field') 'action'
$boundaryRules = @($boundaryLayers.rows)
$ruleActions = [System.Collections.Generic.List[object]]::new()
$ruleExclusionsByTable = @{}
$closureExclusionsByTable = @{}
$clearedRowsByTable = @{}
$catalogByTable = @{}
foreach ($row in $catalog) {
    if (-not $catalogByTable.ContainsKey($row.physical_table)) { $catalogByTable[$row.physical_table] = [System.Collections.Generic.List[object]]::new() }
    $catalogByTable[$row.physical_table].Add($row)
}
for ($ruleIndex = 0; $ruleIndex -lt $boundaryRules.Count; $ruleIndex++) {
    $rule = $boundaryRules[$ruleIndex]
    if ([string]$rule.action -notin @('CLEAR_FIELD','CLEAR_ROW')) { throw "Boundary rule $($ruleIndex+1): action '$($rule.action)' is not supported by script 41." }
    $targetCode = $null
    if ([string]$rule.target -and [string]$rule.target -ne '*') {
        $targetTypes = @(Resolve-ObjectPattern ([string]$rule.target) -AllowMissing)
        if ($targetTypes.Count -eq 0) { continue }
        $targetCode = Get-TypeCode $targetTypes[0]
    }
    $matched = 0
    $ruleTypes = @(Resolve-ObjectPattern ([string]$rule.object) -AllowMissing)
    if ($ruleTypes.Count -eq 0) { continue }
    foreach ($type in $ruleTypes) {
        foreach ($table in $catalogByTable.Keys) {
            $logical = $table.TrimStart('_') -replace 'X[0-9]+$',''
            if ([string]$rule.table_part) {
                if ($logical -notmatch ('^' + $type + '_(?<vt>VT[0-9]+)$')) { continue }
                if ([string]$rule.table_part -ne '*' -and $typeNames[[string]$Matches.vt] -ne [string]$rule.table_part) { continue }
            }
            elseif ($logical -ne $type) { continue }
            foreach ($column in $catalogByTable[$table]) {
                if ($column.kind -notin $refKinds -or $column.field_name -ne [string]$rule.field) { continue }
                $ownerColumn = if ([string]$rule.table_part) { ($catalogByTable[$table] | Where-Object kind -eq 'tabular-owner' | Select-Object -First 1).column } else { '_IDRRef' }
                $ruleActions.Add([pscustomobject]@{
                    rule=$ruleIndex+1; action=[string]$rule.action; object=[string]$rule.object; field=[string]$rule.field
                    physical_table=$table; column=$column.column; type_column=$column.type_column; owner_column=$ownerColumn
                    target_type_code=$targetCode; target_table=$column.target_table
                })
                $match = if ($null -eq $targetCode) { "v.col=N'$($column.column)'" } else { "(v.col=N'$($column.column)' AND v.tc=$targetCode)" }
                if (-not $ruleExclusionsByTable.ContainsKey($table)) {
                    $ruleExclusionsByTable[$table] = [System.Collections.Generic.List[string]]::new()
                    $closureExclusionsByTable[$table] = [System.Collections.Generic.List[string]]::new()
                }
                $ruleExclusionsByTable[$table].Add($match)
                # release=yes: the reference never holds its target in KEEP.
                # Otherwise the rule only resolves references to objects that
                # are deleted for other reasons (foreign-only targets), and the
                # closure still keeps everything else the field points to.
                $closureExclusionsByTable[$table].Add($(if ([string]$rule.release -eq 'yes') { $match } else {
                    "($match AND EXISTS(SELECT 1 FROM #Obj fo WHERE fo.type_code=v.tc AND fo.id=v.id AND fo.has_foreign=1 AND fo.has_keep=0))" }))
                if ([string]$rule.action -eq 'CLEAR_ROW') {
                    # A row that CLEAR_ROW deletes must not hold anything in
                    # KEEP through its other columns either (run 1 kept a
                    # counterparty only through such a row).
                    $columnType = if ($column.type_column) { "CONVERT(int,CONVERT(varbinary(4),src.[$($column.type_column)]))" } else { [string](Get-TypeCode ([string]$column.target_table)) }
                    $rowCondition = "EXISTS(SELECT 1 FROM #Obj fo WHERE fo.type_code=$columnType AND fo.id=src.[$($column.column)] AND fo.has_foreign=1 AND fo.has_keep=0" +
                                    $(if ($null -ne $targetCode) { " AND fo.type_code=$targetCode" } else { '' }) + ')'
                    if (-not $clearedRowsByTable.ContainsKey($table)) { $clearedRowsByTable[$table] = [System.Collections.Generic.List[string]]::new() }
                    $clearedRowsByTable[$table].Add($rowCondition)
                }
                $matched++
            }
        }
    }
    if ($matched -eq 0) { Write-Warning "Boundary rule $($ruleIndex+1) ($($rule.object) / $($rule.field)) matched no column." }
}
function Get-RuleFilter([string]$Table) {
    # Cascade and CLEAR_ROW scans: ruled references are cleared by script 42.
    if (-not $ruleExclusionsByTable.ContainsKey($Table)) { return '' }
    return ' AND NOT (' + ($ruleExclusionsByTable[$Table] -join ' OR ') + ')'
}
function Get-ClosureFilter([string]$Table) {
    # KEEP closure scans: see release above; rows removed by CLEAR_ROW are skipped.
    $filter = ''
    if ($closureExclusionsByTable.ContainsKey($Table)) { $filter += ' AND NOT (' + ($closureExclusionsByTable[$Table] -join ' OR ') + ')' }
    if ($clearedRowsByTable.ContainsKey($Table)) { $filter += ' AND NOT (' + ($clearedRowsByTable[$Table] -join ' OR ') + ')' }
    return $filter
}

# ---------------------------------------------------------------- database
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$connection = [System.Data.SqlClient.SqlConnection]::new(
    "Server=$Server;Database=$Database;Integrated Security=True;Encrypt=False;Application Name=Codex-KeepClosure")
$started = Get-Date
try {
    $connection.Open()
    $identity = Invoke-Table $connection 'SELECT DB_NAME() AS database_name, d.state_desc, d.recovery_model_desc FROM sys.databases d WHERE d.database_id=DB_ID();' 30
    if ([string]$identity.Rows[0].database_name -ne $Database) { throw "Connected to an unexpected database. Expected $Database." }
    if ([string]$identity.Rows[0].state_desc -ne 'ONLINE') { throw 'Database is not ONLINE.' }

    $keptRows = Invoke-Table $connection "SELECT sys.fn_varbintohexstr(_IDRRef) AS id_hex, CONVERT(int,CONVERT(varbinary(1),_Marked)) AS marked FROM dbo.$(Quote-SqlIdentifier $orgTable) WHERE $($organizationFilter.sql);" 60
    if ($keptRows.Rows.Count -ne $organizationFilter.expected) { throw "Expected $($organizationFilter.expected) kept organization(s) '$($organizationFilter.label)'; found $($keptRows.Rows.Count)." }
    foreach ($row in $keptRows.Rows) { if ([int]$row.marked -ne 0) { throw "Kept organization $($row.id_hex) is marked for deletion." } }
    $keepIds = @($keptRows.Rows | ForEach-Object { [string]$_.id_hex } | Sort-Object)
    Invoke-NonQuery $connection ("CREATE TABLE #KeepOrg(id binary(16) NOT NULL PRIMARY KEY); INSERT #KeepOrg(id) VALUES " + (@($keepIds | ForEach-Object { "(CONVERT(binary(16),'$_',1))" }) -join ',') + ';')
    $organizationCounts = Invoke-Table $connection "SELECT COUNT_BIG(*) AS total, SUM(CASE WHEN _Marked=0x01 THEN 1 ELSE 0 END) AS marked FROM dbo.$(Quote-SqlIdentifier $orgTable);" 60

    $predefinedTables = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in (Invoke-Table $connection "SELECT t.name FROM sys.tables t JOIN sys.columns c ON c.object_id=t.object_id WHERE c.name=N'_PredefinedID';" 60).Rows) {
        [void]$predefinedTables.Add([string]$row.name)
    }

    Invoke-NonQuery $connection @"
CREATE TABLE #Obj(type_code int NOT NULL, id binary(16) NOT NULL, marked bit NOT NULL DEFAULT 0, predefined bit NOT NULL DEFAULT 0,
                  has_keep bit NOT NULL DEFAULT 0, has_foreign bit NOT NULL DEFAULT 0,
                  PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #Keep(type_code int NOT NULL, id binary(16) NOT NULL, round int NOT NULL, reason nvarchar(64) NOT NULL,
                   PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #Delta(type_code int NOT NULL, id binary(16) NOT NULL, PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #Cand(type_code int NOT NULL, id binary(16) NOT NULL, source_table nvarchar(128) NOT NULL, source_column nvarchar(128) NOT NULL,
                   source_type_code int NULL, source_id binary(16) NULL,
                   PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #Boundary(type_code int NOT NULL, id binary(16) NOT NULL, target_class nvarchar(32) NOT NULL,
                       source_table nvarchar(128) NOT NULL, source_column nvarchar(128) NOT NULL,
                       source_type_code int NULL, source_id binary(16) NULL, round int NOT NULL);
CREATE TABLE #Missing(type_code int NOT NULL, id binary(16) NOT NULL, source_table nvarchar(128) NOT NULL, source_column nvarchar(128) NOT NULL,
                      PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #Types(type_code int NOT NULL PRIMARY KEY, type_table nvarchar(128) NOT NULL);
"@
    $typeRows = @($objectTypeByCode.GetEnumerator() | ForEach-Object { "($($_.Key),N'$($_.Value)')" })
    for ($i = 0; $i -lt $typeRows.Count; $i += 900) {
        Invoke-NonQuery $connection ("INSERT #Types(type_code,type_table) VALUES " + ($typeRows[$i..([Math]::Min($i+899,$typeRows.Count-1))] -join ','))
    }

    # 1. Object inventory and organization classification.
    $statements = [System.Collections.Generic.List[string]]::new()
    foreach ($type in ($objectTablesByType.Keys | Sort-Object)) {
        $code = Get-TypeCode $type
        foreach ($table in $objectTablesByType[$type]) {
            $predefined = if ($predefinedTables.Contains($table)) { "CASE WHEN src.[_PredefinedID]<>$zeroId THEN 1 ELSE 0 END" } else { '0' }
            $statements.Add("INSERT #Obj(type_code,id,marked,predefined) SELECT $code,src.[_IDRRef],CASE WHEN src.[_Marked]=0x01 THEN 1 ELSE 0 END,$predefined FROM dbo.$(Quote-SqlIdentifier $table) src;")
        }
    }
    Invoke-Batched $connection $statements 'Object inventory'
    $statements.Clear()
    foreach ($type in ($objectTablesByType.Keys | Sort-Object)) {
        $code = Get-TypeCode $type
        foreach ($table in $objectTablesByType[$type]) {
            if (-not $orgPredicatesByTable.ContainsKey($table)) { continue }
            $keepPredicate = (@($orgPredicatesByTable[$table] | ForEach-Object keep) -join ' OR ')
            $foreignPredicate = (@($orgPredicatesByTable[$table] | ForEach-Object foreign) -join ' OR ')
            $statements.Add("UPDATE o SET has_keep=1 FROM #Obj o JOIN dbo.$(Quote-SqlIdentifier $table) src ON o.type_code=$code AND o.id=src.[_IDRRef] WHERE $keepPredicate;")
            $statements.Add("UPDATE o SET has_foreign=1 FROM #Obj o JOIN dbo.$(Quote-SqlIdentifier $table) src ON o.type_code=$code AND o.id=src.[_IDRRef] WHERE $foreignPredicate;")
        }
    }
    Invoke-Batched $connection $statements 'Organization classification'
    # The organization catalog itself: only the kept rows are "ours".
    Invoke-NonQuery $connection "UPDATE o SET has_keep=CASE WHEN k.id IS NULL THEN 0 ELSE 1 END, has_foreign=CASE WHEN k.id IS NULL THEN 1 ELSE 0 END FROM #Obj o LEFT JOIN #KeepOrg k ON k.id=o.id WHERE o.type_code=$orgCode;"

    # 2. Seeds.
    Invoke-NonQuery $connection @"
INSERT #Keep(type_code,id,round,reason) SELECT type_code,id,0,N'kept-organization' FROM #Obj WHERE type_code=$orgCode AND has_keep=1;
INSERT #Keep(type_code,id,round,reason) SELECT type_code,id,0,N'organization-field' FROM #Obj WHERE has_keep=1;
INSERT #Keep(type_code,id,round,reason) SELECT type_code,id,0,N'predefined' FROM #Obj WHERE predefined=1;
"@
    foreach ($root in $keepRoots) {
        # KEEP_ROOT condition: "<Реквизит><>0" keeps only rows whose field is
        # filled (e.g. users bound to an infobase account).
        $rootCode = Get-TypeCode $root.type
        foreach ($table in $objectTablesByType[$root.type]) {
            $where = ''
            if ($root.condition) {
                if ($root.condition -notmatch '^(?<field>[^<>=]+)<>0$') { throw "Unsupported KEEP_ROOT condition: $($root.condition)" }
                $fieldCodes = @($fieldNameByCode.GetEnumerator() | Where-Object { $_.Value -eq $Matches.field } | ForEach-Object { '_' + $_.Name })
                $column = @((Invoke-Table $connection "SELECT name FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.$table');" 60).Rows | ForEach-Object { [string]$_.name } | Where-Object { $_ -in $fieldCodes }) | Select-Object -First 1
                if (-not $column) { continue }
                $where = " WHERE src.$(Quote-SqlIdentifier $column) IS NOT NULL AND src.$(Quote-SqlIdentifier $column)<>0x" + ('0' * 32)
            }
            Invoke-NonQuery $connection "INSERT #Keep(type_code,id,round,reason) SELECT $rootCode,src.[_IDRRef],0,N'keep-root' FROM dbo.$(Quote-SqlIdentifier $table) src$where;"
        }
    }
    $statements.Clear()
    $registerRows = [System.Collections.Generic.List[object]]::new()
    foreach ($table in $registerOrgTables) {
        $keepPredicate = (@($orgPredicatesByTable[$table] | ForEach-Object keep) -join ' OR ')
        $recorder = $recorderByTable[$table]
        if ($recorder) {
            $recorderId = Quote-SqlIdentifier ([string]$recorder.column)
            $recorderType = if ($recorder.type_column) { "CONVERT(int,CONVERT(varbinary(4),src.$(Quote-SqlIdentifier ([string]$recorder.type_column))))" } else { [string](Get-TypeCode ([string]$recorder.target_table)) }
            $statements.Add("INSERT #Keep(type_code,id,round,reason) SELECT DISTINCT $recorderType,src.$recorderId,0,N'recorder-of-organization-register-row' FROM dbo.$(Quote-SqlIdentifier $table) src WHERE ($keepPredicate) AND src.$recorderId<>$zeroId;")
        }
        if ($refValuesByTable.ContainsKey($table)) {
            $statements.Add("INSERT #Cand(type_code,id,source_table,source_column) SELECT v.tc,v.id,N'$table',v.col FROM dbo.$(Quote-SqlIdentifier $table) src CROSS APPLY (VALUES $($refValuesByTable[$table] -join ',')) v(tc,id,col) WHERE ($keepPredicate) AND v.id IS NOT NULL AND v.id<>$zeroId$(Get-ClosureFilter $table);")
        }
        foreach ($location in $orgPredicatesByTable[$table]) {
            $registerRows.Add([pscustomobject]@{ physical_table=$table; id_column=$location.column; type_column=$location.type_column; keep=$location.keep; has_recorder=[int][bool]$recorder })
        }
    }
    foreach ($table in $constantTables) {
        $statements.Add("INSERT #Cand(type_code,id,source_table,source_column) SELECT v.tc,v.id,N'$table',v.col FROM dbo.$(Quote-SqlIdentifier $table) src CROSS APPLY (VALUES $($refValuesByTable[$table] -join ',')) v(tc,id,col) WHERE v.id IS NOT NULL AND v.id<>$zeroId$(Get-ClosureFilter $table);")
    }
    Invoke-Batched $connection $statements 'Seed from organization registers and constants'

    # Protected register-row manifest in the format of script 27 / script 11.
    $registerManifest = [System.Collections.Generic.List[object]]::new()
    foreach ($location in $registerRows) {
        $count = [int64](Invoke-Scalar $connection "SELECT COUNT_BIG(*) FROM dbo.$(Quote-SqlIdentifier $location.physical_table) src WHERE $($location.keep);")
        $logical = $location.physical_table.TrimStart('_') -replace 'X[0-9]+$',''
        $registerManifest.Add([pscustomobject]@{
            type_table=$logical; type_code=(Get-TypeCode $logical); object_name=$typeNames[$logical]
            physical_table=$location.physical_table; id_column=$location.id_column; type_column=$location.type_column
            matched_rows=$count; has_recorder=$location.has_recorder
        })
    }
    $registerManifestPath = Join-Path $OutputDirectory 'organization-register-rows.csv'
    $registerManifest | Export-Csv -LiteralPath $registerManifestPath -NoTypeInformation -Encoding UTF8

    # 3. Closure.  Classification of a candidate reference:
    #    absent object         -> #Missing (baseline, not a decision)
    #    other organization    -> boundary
    #    foreign-only object   -> boundary
    #    anything else         -> KEEP (shared)
    $classifySql = @"
DELETE c FROM #Cand c WHERE EXISTS(SELECT 1 FROM #Keep k WHERE k.type_code=c.type_code AND k.id=c.id);
INSERT #Missing(type_code,id,source_table,source_column)
SELECT c.type_code,c.id,c.source_table,c.source_column FROM #Cand c
WHERE EXISTS(SELECT 1 FROM #Types t WHERE t.type_code=c.type_code)
  AND NOT EXISTS(SELECT 1 FROM #Obj o WHERE o.type_code=c.type_code AND o.id=c.id);
INSERT #Boundary(type_code,id,target_class,source_table,source_column,source_type_code,source_id,round)
SELECT c.type_code,c.id,CASE WHEN c.type_code=$orgCode THEN N'foreign-organization' WHEN t.type_table LIKE N'Document%' THEN N'foreign-document' ELSE N'foreign-owned' END,
       c.source_table,c.source_column,c.source_type_code,c.source_id,@Round
FROM #Cand c JOIN #Obj o ON o.type_code=c.type_code AND o.id=c.id JOIN #Types t ON t.type_code=c.type_code
WHERE o.has_foreign=1 AND o.has_keep=0;
TRUNCATE TABLE #Delta;
INSERT #Delta(type_code,id)
SELECT c.type_code,c.id FROM #Cand c JOIN #Obj o ON o.type_code=c.type_code AND o.id=c.id
WHERE NOT (o.has_foreign=1 AND o.has_keep=0);
INSERT #Keep(type_code,id,round,reason) SELECT type_code,id,@Round,N'referenced-by-keep' FROM #Delta;
TRUNCATE TABLE #Cand;
"@
    Invoke-NonQuery $connection ("DECLARE @Round int=0;`n" + $classifySql)
    Invoke-NonQuery $connection 'INSERT #Delta(type_code,id) SELECT type_code,id FROM #Keep;'

    $rounds = [System.Collections.Generic.List[object]]::new()
    for ($round = 1; $round -le $MaxRounds; $round++) {
        $deltaTypes = Invoke-Table $connection 'SELECT d.type_code, COUNT_BIG(*) AS objects FROM #Delta d GROUP BY d.type_code;'
        $deltaObjects = [int64](($deltaTypes.Rows | Measure-Object -Property objects -Sum).Sum)
        if ($deltaObjects -eq 0) { break }
        $statements.Clear()
        foreach ($typeRow in $deltaTypes.Rows) {
            $code = [int]$typeRow.type_code
            $type = $objectTypeByCode[$code]
            foreach ($table in $objectTablesByType[$type]) {
                if (-not $refValuesByTable.ContainsKey($table)) { continue }
                $statements.Add("INSERT #Cand(type_code,id,source_table,source_column,source_type_code,source_id) SELECT v.tc,v.id,N'$table',v.col,$code,d.id FROM #Delta d JOIN dbo.$(Quote-SqlIdentifier $table) src ON src.[_IDRRef]=d.id CROSS APPLY (VALUES $($refValuesByTable[$table] -join ',')) v(tc,id,col) WHERE d.type_code=$code AND v.id IS NOT NULL AND v.id<>$zeroId$(Get-ClosureFilter $table);")
            }
            if (-not $dependentSourcesByType.ContainsKey($type)) { continue }
            foreach ($source in $dependentSourcesByType[$type]) {
                $typeFilter = if ($source.type_column) { " AND src.$(Quote-SqlIdentifier $source.type_column)=CONVERT(binary(4),$code)" } else { '' }
                if ($null -ne $source.self_code) {
                    $statements.Add("INSERT #Cand(type_code,id,source_table,source_column,source_type_code,source_id) SELECT $($source.self_code),src.[_IDRRef],N'$($source.table)',N'$($source.key)',$code,d.id FROM #Delta d JOIN dbo.$(Quote-SqlIdentifier $source.table) src ON src.$(Quote-SqlIdentifier $source.key)=d.id$typeFilter WHERE d.type_code=$code;")
                }
                if (-not $refValuesByTable.ContainsKey($source.table)) { continue }
                $statements.Add("INSERT #Cand(type_code,id,source_table,source_column,source_type_code,source_id) SELECT v.tc,v.id,N'$($source.table)',v.col,$code,d.id FROM #Delta d JOIN dbo.$(Quote-SqlIdentifier $source.table) src ON src.$(Quote-SqlIdentifier $source.key)=d.id$typeFilter CROSS APPLY (VALUES $($refValuesByTable[$source.table] -join ',')) v(tc,id,col) WHERE d.type_code=$code AND v.id IS NOT NULL AND v.id<>$zeroId$(Get-ClosureFilter $source.table);")
            }
        }
        $roundStarted = Get-Date
        Invoke-Batched $connection $statements "Round $round ($deltaObjects new KEEP objects)"
        Invoke-NonQuery $connection ("DECLARE @Round int=$round;`n" + $classifySql)
        $added = [int64](Invoke-Scalar $connection 'SELECT COUNT_BIG(*) FROM #Delta;')
        $rounds.Add([pscustomobject]@{ round=$round; scanned_objects=$deltaObjects; source_statements=$statements.Count; added_objects=$added; seconds=[Math]::Round(((Get-Date)-$roundStarted).TotalSeconds,1) })
        Write-Host ("Round {0}: scanned {1}, added {2}" -f $round, $deltaObjects, $added)
    }
    if ([int64](Invoke-Scalar $connection 'SELECT COUNT_BIG(*) FROM #Delta;') -ne 0) { throw "KEEP closure did not reach a fixed point in $MaxRounds rounds." }

    # 4. DELETE as the complement.
    Invoke-NonQuery $connection @"
CREATE TABLE #Delete(type_code int NOT NULL, id binary(16) NOT NULL, reason nvarchar(64) NOT NULL, PRIMARY KEY(type_code,id));
INSERT #Delete(type_code,id,reason)
SELECT o.type_code,o.id,
       CASE WHEN o.type_code=$orgCode THEN N'other-organization'
            WHEN t.type_table LIKE N'Document%' AND o.has_foreign=1 THEN N'document-of-other-organization'
            WHEN t.type_table LIKE N'Document%' THEN N'document-without-kept-organization'
            WHEN o.has_foreign=1 AND o.has_keep=0 THEN N'owned-by-other-organization'
            ELSE N'complement-catalog' END
FROM #Obj o JOIN #Types t ON t.type_code=o.type_code
WHERE NOT EXISTS(SELECT 1 FROM #Keep k WHERE k.type_code=o.type_code AND k.id=o.id)
  AND o.predefined=0
  AND (o.type_code=$orgCode OR t.type_table LIKE N'Document%' OR (o.has_foreign=1 AND o.has_keep=0) OR o.type_code IN ($complementList));
"@

    # 5. Cascade over objects outside KEEP and DELETE.
    $protectedCascadeCodes = @(@($typeNames.GetEnumerator() | Where-Object { $_.Value -in $CascadeProtectedTypeNames -and $objectTablesByType.ContainsKey($_.Name) } | ForEach-Object { Get-TypeCode $_.Name }) + $protectCodes | Sort-Object -Unique)
    $protectedCascadeList = if ($protectedCascadeCodes.Count) { $protectedCascadeCodes -join ',' } else { '-1' }
    Invoke-NonQuery $connection @'
CREATE TABLE #Cascade(type_code int NOT NULL, id binary(16) NOT NULL, round int NOT NULL, source_table nvarchar(128) NOT NULL, source_column nvarchar(128) NOT NULL,
                      target_type_code int NOT NULL, target_id binary(16) NOT NULL, PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #CascadeHit(type_code int NOT NULL, id binary(16) NOT NULL, source_table nvarchar(128) NOT NULL, source_column nvarchar(128) NOT NULL,
                         target_type_code int NOT NULL, target_id binary(16) NOT NULL, PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #CascadeBlocked(type_code int NOT NULL, id binary(16) NOT NULL, source_table nvarchar(128) NOT NULL, source_column nvarchar(128) NOT NULL,
                             target_type_code int NOT NULL, target_id binary(16) NOT NULL, PRIMARY KEY(type_code,id) WITH (IGNORE_DUP_KEY=ON));
CREATE TABLE #ClearRow(owner_type_code int NOT NULL, owner_id binary(16) NOT NULL, source_table nvarchar(128) NOT NULL, source_column nvarchar(128) NOT NULL,
                       target_type_code int NOT NULL, target_id binary(16) NOT NULL, rows_ bigint NOT NULL);
CREATE TABLE #Gone(type_code int NOT NULL, id binary(16) NOT NULL, PRIMARY KEY(type_code,id));
CREATE TABLE #GoneDelta(type_code int NOT NULL, id binary(16) NOT NULL, PRIMARY KEY(type_code,id));
INSERT #Gone(type_code,id) SELECT type_code,id FROM #Delete;
INSERT #GoneDelta(type_code,id) SELECT type_code,id FROM #Delete;
'@
    $cascadeRounds = [System.Collections.Generic.List[object]]::new()
    for ($round = 1; $round -le $MaxRounds; $round++) {
        $neutralTypes = Invoke-Table $connection @'
SELECT o.type_code FROM #Obj o
WHERE NOT EXISTS(SELECT 1 FROM #Keep k WHERE k.type_code=o.type_code AND k.id=o.id)
  AND NOT EXISTS(SELECT 1 FROM #Gone g WHERE g.type_code=o.type_code AND g.id=o.id)
  AND NOT EXISTS(SELECT 1 FROM #CascadeBlocked b WHERE b.type_code=o.type_code AND b.id=o.id)
GROUP BY o.type_code;
'@
        $statements.Clear()
        foreach ($typeRow in $neutralTypes.Rows) {
            $code = [int]$typeRow.type_code
            $type = $objectTypeByCode[$code]
            # Only a header reference cascades the whole object.  A reference
            # from one tabular row (a per-organization setting of a
            # nomenclature group, say) must not delete the owner and, through
            # it, everything that uses the owner; such rows are collected as
            # CLEAR_ROW candidates instead.
            foreach ($table in $objectTablesByType[$type]) {
                if (-not $refValuesByTable.ContainsKey($table)) { continue }
                $statements.Add("INSERT #CascadeHit(type_code,id,source_table,source_column,target_type_code,target_id) SELECT $code,src.[_IDRRef],N'$table',v.col,v.tc,v.id FROM dbo.$(Quote-SqlIdentifier $table) src CROSS APPLY (VALUES $($refValuesByTable[$table] -join ',')) v(tc,id,col) JOIN #GoneDelta g ON g.type_code=v.tc AND g.id=v.id WHERE NOT EXISTS(SELECT 1 FROM #Keep k WHERE k.type_code=$code AND k.id=src.[_IDRRef]) AND NOT EXISTS(SELECT 1 FROM #Gone x WHERE x.type_code=$code AND x.id=src.[_IDRRef])$(Get-RuleFilter $table);")
            }
            if ($dependentSourcesByType.ContainsKey($type)) {
                foreach ($source in $dependentSourcesByType[$type]) {
                    if ($source.type_column -or -not $refValuesByTable.ContainsKey($source.table)) { continue }
                    $key = Quote-SqlIdentifier $source.key
                    $statements.Add("INSERT #ClearRow(owner_type_code,owner_id,source_table,source_column,target_type_code,target_id,rows_) SELECT $code,src.$key,N'$($source.table)',v.col,v.tc,v.id,COUNT_BIG(*) FROM dbo.$(Quote-SqlIdentifier $source.table) src CROSS APPLY (VALUES $($refValuesByTable[$source.table] -join ',')) v(tc,id,col) JOIN #GoneDelta g ON g.type_code=v.tc AND g.id=v.id WHERE NOT EXISTS(SELECT 1 FROM #Gone x WHERE x.type_code=$code AND x.id=src.$key)$(Get-RuleFilter $source.table) GROUP BY src.$key,v.col,v.tc,v.id;")
                }
            }
        }
        Invoke-Batched $connection $statements "Cascade round $round"
        $added = [int64](Invoke-Scalar $connection @"
INSERT #CascadeBlocked SELECT h.* FROM #CascadeHit h WHERE h.type_code IN ($protectedCascadeList);
DELETE FROM #CascadeHit WHERE type_code IN ($protectedCascadeList);
TRUNCATE TABLE #GoneDelta;
INSERT #GoneDelta(type_code,id) SELECT h.type_code,h.id FROM #CascadeHit h WHERE NOT EXISTS(SELECT 1 FROM #Gone g WHERE g.type_code=h.type_code AND g.id=h.id);
INSERT #Gone(type_code,id) SELECT type_code,id FROM #GoneDelta;
INSERT #Cascade(type_code,id,round,source_table,source_column,target_type_code,target_id)
SELECT h.type_code,h.id,$round,h.source_table,h.source_column,h.target_type_code,h.target_id FROM #CascadeHit h JOIN #GoneDelta d ON d.type_code=h.type_code AND d.id=h.id;
TRUNCATE TABLE #CascadeHit;
SELECT COUNT_BIG(*) FROM #GoneDelta;
"@)
        $cascadeRounds.Add([pscustomobject]@{ round=$round; source_statements=$statements.Count; added_objects=$added })
        Write-Host ("Cascade round {0}: added {1}" -f $round, $added)
        if ($added -eq 0) { break }
    }

    $cascadePath = Join-Path $OutputDirectory 'cascade-manifest.csv'
    $cascadeEvidencePath = Join-Path $OutputDirectory 'cascade-by-field.csv'
    $cascadeBlockedPath = Join-Path $OutputDirectory 'cascade-blocked.csv'
    $deleteWithCascadePath = Join-Path $OutputDirectory 'delete-manifest-with-cascade.csv'
    $cascadeRows = Export-QueryCsv $connection @'
SELECT t.type_table, c.type_code, sys.fn_varbintohexstr(c.id) AS id_hex, c.round AS graph_depth, N'cascade-references-deleted' AS inclusion_reason
FROM #Cascade c JOIN #Types t ON t.type_code=c.type_code ORDER BY t.type_table, id_hex;
'@ $cascadePath
    [void](Export-QueryCsv $connection @'
SELECT t.type_table, c.source_table, c.source_column, tt.type_table AS target_type_table, MIN(c.round) AS first_round, COUNT_BIG(*) AS objects
FROM #Cascade c JOIN #Types t ON t.type_code=c.type_code JOIN #Types tt ON tt.type_code=c.target_type_code
GROUP BY t.type_table, c.source_table, c.source_column, tt.type_table ORDER BY objects DESC;
'@ $cascadeEvidencePath)
    $cascadeBlockedRows = Export-QueryCsv $connection @'
SELECT t.type_table, sys.fn_varbintohexstr(b.id) AS id_hex, b.source_table, b.source_column, tt.type_table AS target_type_table, sys.fn_varbintohexstr(b.target_id) AS target_id_hex
FROM #CascadeBlocked b JOIN #Types t ON t.type_code=b.type_code JOIN #Types tt ON tt.type_code=b.target_type_code ORDER BY t.type_table, id_hex;
'@ $cascadeBlockedPath
    $clearRowPath = Join-Path $OutputDirectory 'clear-row-candidates.csv'
    $clearRowCandidates = Export-QueryCsv $connection @'
SELECT ot.type_table AS owner_type_table, sys.fn_varbintohexstr(c.owner_id) AS owner_id_hex,
       CASE WHEN EXISTS(SELECT 1 FROM #Keep k WHERE k.type_code=c.owner_type_code AND k.id=c.owner_id) THEN N'keep' ELSE N'untouched' END AS owner_state,
       c.source_table, c.source_column, tt.type_table AS target_type_table, sys.fn_varbintohexstr(c.target_id) AS target_id_hex, SUM(c.rows_) AS rows_
FROM #ClearRow c JOIN #Types ot ON ot.type_code=c.owner_type_code JOIN #Types tt ON tt.type_code=c.target_type_code
WHERE NOT EXISTS(SELECT 1 FROM #Gone g WHERE g.type_code=c.owner_type_code AND g.id=c.owner_id)
GROUP BY ot.type_table, c.owner_type_code, c.owner_id, c.source_table, c.source_column, tt.type_table, c.target_id
ORDER BY c.source_table, owner_id_hex;
'@ $clearRowPath
    [void](Export-QueryCsv $connection @'
SELECT t.type_table, g.type_code, sys.fn_varbintohexstr(g.id) AS id_hex, 0 AS graph_depth,
       COALESCE(d.reason, N'cascade-references-deleted') AS inclusion_reason
FROM #Gone g JOIN #Types t ON t.type_code=g.type_code LEFT JOIN #Delete d ON d.type_code=g.type_code AND d.id=g.id
ORDER BY t.type_table, id_hex;
'@ $deleteWithCascadePath)

    # Column-level actions of the boundary rules, executed by script 42 against
    # the frozen delete-with-cascade manifest before script 11 runs.
    $ruleActionsPath = Join-Path $OutputDirectory 'rule-actions.csv'
    # Sorted: the rules walk hashtable keys, whose order changes between
    # processes, and the file hash is a frozen input of script 42.
    $ruleActions | Sort-Object rule,physical_table,column | Select-Object rule,action,object,field,physical_table,column,type_column,owner_column,target_type_code,target_table |
        Export-Csv -LiteralPath $ruleActionsPath -NoTypeInformation -Encoding UTF8

    $keepPath = Join-Path $OutputDirectory 'keep-manifest.csv'
    $deletePath = Join-Path $OutputDirectory 'delete-manifest.csv'
    $boundaryPath = Join-Path $OutputDirectory 'boundary-references.csv'
    $missingPath = Join-Path $OutputDirectory 'missing-references.csv'
    $keepRows = Export-QueryCsv $connection @'
SELECT t.type_table, k.type_code, sys.fn_varbintohexstr(k.id) AS id_hex, k.round AS graph_depth, k.reason AS inclusion_reason
FROM #Keep k JOIN #Types t ON t.type_code=k.type_code ORDER BY t.type_table, id_hex;
'@ $keepPath
    $deleteRows = Export-QueryCsv $connection @'
SELECT t.type_table, d.type_code, sys.fn_varbintohexstr(d.id) AS id_hex, 0 AS graph_depth, d.reason AS inclusion_reason
FROM #Delete d JOIN #Types t ON t.type_code=d.type_code ORDER BY t.type_table, id_hex;
'@ $deletePath
    $boundaryRows = Export-QueryCsv $connection @'
SELECT b.target_class, t.type_table AS target_type_table, sys.fn_varbintohexstr(b.id) AS target_id_hex,
       b.source_table, b.source_column, st.type_table AS source_owner_type_table, sys.fn_varbintohexstr(b.source_id) AS source_owner_id_hex, b.round
FROM #Boundary b JOIN #Types t ON t.type_code=b.type_code LEFT JOIN #Types st ON st.type_code=b.source_type_code
ORDER BY b.target_class, target_type_table, target_id_hex, b.source_table, b.source_column;
'@ $boundaryPath
    $missingRows = Export-QueryCsv $connection @'
SELECT t.type_table, sys.fn_varbintohexstr(m.id) AS id_hex, m.source_table, m.source_column
FROM #Missing m JOIN #Types t ON t.type_code=m.type_code ORDER BY t.type_table, id_hex;
'@ $missingPath

    $byType = Invoke-Table $connection @'
SELECT t.type_table,
       (SELECT COUNT_BIG(*) FROM #Obj o WHERE o.type_code=t.type_code) AS objects,
       (SELECT COUNT_BIG(*) FROM #Keep k WHERE k.type_code=t.type_code) AS keep_objects,
       (SELECT COUNT_BIG(*) FROM #Delete d WHERE d.type_code=t.type_code) AS delete_objects,
       (SELECT COUNT_BIG(*) FROM #Obj o WHERE o.type_code=t.type_code AND o.has_keep=1) AS with_kept_organization,
       (SELECT COUNT_BIG(*) FROM #Obj o WHERE o.type_code=t.type_code AND o.has_foreign=1 AND o.has_keep=0) AS other_organization_only
FROM #Types t;
'@
    $byTypePath = Join-Path $OutputDirectory 'summary-by-type.csv'
    @($byType.Rows | Where-Object { [int64]$_.objects -gt 0 } | ForEach-Object {
        [pscustomobject]@{
            type_table=$_.type_table; logical_name=$typeNames[[string]$_.type_table]; objects=$_.objects
            keep_objects=$_.keep_objects; delete_objects=$_.delete_objects
            untouched_objects=[int64]$_.objects-[int64]$_.keep_objects-[int64]$_.delete_objects
            with_kept_organization=$_.with_kept_organization; other_organization_only=$_.other_organization_only
        }
    } | Sort-Object type_table) | Export-Csv -LiteralPath $byTypePath -NoTypeInformation -Encoding UTF8

    $counts = Invoke-Table $connection @"
SELECT (SELECT COUNT_BIG(*) FROM #Obj) AS objects,
       (SELECT COUNT_BIG(*) FROM #Keep) AS keep_objects,
       (SELECT COUNT_BIG(*) FROM #Delete) AS delete_objects,
       (SELECT COUNT_BIG(*) FROM #Keep k JOIN #Delete d ON d.type_code=k.type_code AND d.id=k.id) AS keep_delete_intersection,
       (SELECT COUNT_BIG(*) FROM #Delete WHERE type_code=$orgCode) AS delete_organizations,
       (SELECT COUNT_BIG(*) FROM #Delete d JOIN #Obj o ON o.type_code=d.type_code AND o.id=d.id WHERE d.type_code=$orgCode AND o.marked=0) AS delete_unmarked_organizations,
       (SELECT COUNT(DISTINCT CONCAT(type_code,'|',sys.fn_varbintohexstr(id))) FROM #Boundary) AS boundary_targets,
       (SELECT COUNT_BIG(*) FROM #Missing) AS missing_targets;
"@
    $reasons = Invoke-Table $connection 'SELECT reason, COUNT_BIG(*) AS objects FROM #Keep GROUP BY reason UNION ALL SELECT N''delete:''+reason, COUNT_BIG(*) FROM #Delete GROUP BY reason;'
    $boundaryClasses = Invoke-Table $connection 'SELECT target_class, COUNT(DISTINCT CONCAT(type_code,''|'',sys.fn_varbintohexstr(id))) AS targets, COUNT_BIG(*) AS references_ FROM #Boundary GROUP BY target_class;'
    $c = $counts.Rows[0]
    $summary = [ordered]@{
        generated_at_local=(Get-Date).ToString('s'); server=$Server; database=$Database; read_only=$true
        organization_code=$organizationFilter.label; organization_type=$orgType
        organization_id_hex=($keepIds -join ','); organization_ids=@($keepIds)
        organizations_total=[int64]$organizationCounts.Rows[0].total; organizations_marked=[int64]$organizationCounts.Rows[0].marked
        catalog_fixed_edges_sha256=$catalogSummary.fixed_edges_sha256
        catalog_policy=(@($catalogLayers.files | ForEach-Object { $_.path }) -join ';')
        catalog_policy_sha256=(@($catalogLayers.files | ForEach-Object { $_.sha256 }) -join ';')
        catalog_policy_layers=@($catalogLayers.files)
        complement_catalogs=@($catalogPolicyRows | Where-Object role -eq 'COMPLEMENT' | ForEach-Object object)
        extension_complement_skipped=@($extensionComplementSkipped)
        keep_root_catalogs=@($catalogPolicyRows | Where-Object role -eq 'KEEP_ROOT' | ForEach-Object object)
        protected_types=$protectCodes.Count; complement_types=$complementCodes.Count
        boundary_rules=(@($boundaryLayers.files | ForEach-Object { $_.path }) -join ';')
        boundary_rules_sha256=(@($boundaryLayers.files | ForEach-Object { $_.sha256 }) -join ';')
        boundary_rule_layers=@($boundaryLayers.files)
        skipped_policy_objects=@($skippedPolicyObjects)
        rule_actions=$ruleActionsPath; rule_action_columns=$ruleActions.Count
        rule_actions_sha256=(Get-FileHash -LiteralPath $ruleActionsPath -Algorithm SHA256).Hash.ToLowerInvariant()
        objects=[int64]$c.objects; keep_objects=[int64]$c.keep_objects; delete_objects=[int64]$c.delete_objects
        keep_delete_intersection=[int64]$c.keep_delete_intersection
        delete_organizations=[int64]$c.delete_organizations; delete_unmarked_organizations=[int64]$c.delete_unmarked_organizations
        boundary_targets=[int64]$c.boundary_targets; missing_reference_targets=[int64]$c.missing_targets
        reasons=[ordered]@{}; boundary_classes=@($boundaryClasses.Rows | ForEach-Object { [ordered]@{ class=$_.target_class; targets=$_.targets; references=$_.references_ } })
        rounds=@($rounds)
        cascade_objects=$cascadeRows; cascade_blocked_objects=$cascadeBlockedRows; cascade_rounds=@($cascadeRounds)
        clear_row_candidates=$clearRowCandidates; clear_row_candidates_csv=$clearRowPath
        cascade_protected_types=@($CascadeProtectedTypeNames)
        gate=if ([int64]$c.keep_delete_intersection -eq 0 -and [int64]$c.boundary_targets -eq 0 -and [int64]$c.delete_unmarked_organizations -eq 0 -and $cascadeBlockedRows -eq 0 -and $clearRowCandidates -eq 0) { 'PASS' } else { 'BLOCK' }
        keep_manifest=$keepPath; keep_manifest_sha256=(Get-FileHash -LiteralPath $keepPath -Algorithm SHA256).Hash.ToLowerInvariant()
        delete_manifest=$deletePath; delete_manifest_sha256=(Get-FileHash -LiteralPath $deletePath -Algorithm SHA256).Hash.ToLowerInvariant()
        delete_manifest_with_cascade=$deleteWithCascadePath; delete_manifest_with_cascade_sha256=(Get-FileHash -LiteralPath $deleteWithCascadePath -Algorithm SHA256).Hash.ToLowerInvariant()
        cascade_manifest=$cascadePath; cascade_by_field=$cascadeEvidencePath; cascade_blocked=$cascadeBlockedPath
        organization_register_rows=$registerManifestPath; organization_register_rows_sha256=(Get-FileHash -LiteralPath $registerManifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
        boundary_references=$boundaryPath; missing_references=$missingPath; summary_by_type=$byTypePath
        duration_seconds=[Math]::Round(((Get-Date)-$started).TotalSeconds,1)
    }
    foreach ($row in $reasons.Rows) { $summary.reasons[[string]$row.reason] = [int64]$row.objects }
    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
    $summary | ConvertTo-Json -Depth 6
}
finally { $connection.Dispose() }
