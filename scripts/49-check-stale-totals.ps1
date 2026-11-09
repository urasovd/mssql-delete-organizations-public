[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$Database,
    [Parameter(Mandatory)] [string]$CatalogDirectory,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-stale-totals'))
)
<#
    Read-only check after the totals recalculation: rows of derived tables
    (register totals, accounting balances, slices) whose organization column
    points to an organization that no longer exists.

    Every such row is a stale total of a deleted organization.  The registers
    listed in stale-totals.csv are the ones to pass to
    47 -RecalcTotals -TotalsRegister.  Only fixed-type organization columns
    are checked; composite columns are left to TiI.
#>
$ErrorActionPreference = 'Stop'

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

$organizationTable = @(Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'type-names.csv') |
    Where-Object { $_.logical_name -eq 'Организации' -and $_.type_table -match '^Reference[0-9]+$' } |
    Select-Object -ExpandProperty type_table)
if ($organizationTable.Count -ne 1) { throw "Справочник.Организации was not found unambiguously in type-names.csv: $($organizationTable -join ', ')" }
$organizationTable = $organizationTable[0]

$columns = @(Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'reference-catalog.csv') |
    Where-Object { $_.table_role -eq 'derived' -and $_.kind -eq 'fixed' -and $_.target_table -eq $organizationTable })

$connection = [System.Data.SqlClient.SqlConnection]::new(
    "Server=$Server;Database=$Database;Integrated Security=True;Encrypt=False;Application Name=Codex-StaleTotals")
$connection.Open()
try {
    $existing = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $command = $connection.CreateCommand()
    $command.CommandText = 'SELECT name FROM sys.tables'
    $reader = $command.ExecuteReader()
    while ($reader.Read()) { [void]$existing.Add($reader.GetString(0)) }
    $reader.Close()

    $rows = foreach ($column in $columns) {
        if (-not $existing.Contains($column.physical_table)) { continue }
        $command = $connection.CreateCommand()
        $command.CommandTimeout = 0
        $command.CommandText = "SELECT COUNT_BIG(*) FROM [$($column.physical_table)] t WITH (NOLOCK) " +
            "WHERE t.[$($column.column)] <> 0x00000000000000000000000000000000 " +
            "AND NOT EXISTS (SELECT 1 FROM [_$organizationTable] o WHERE o._IDRRef = t.[$($column.column)])"
        $stale = [long]$command.ExecuteScalar()
        [pscustomobject]@{ register=$column.table_name; physical_table=$column.physical_table; column=$column.column; stale_rows=$stale }
    }
} finally { $connection.Close() }

$rows = @($rows | Sort-Object register, physical_table, column)
$path = Join-Path $OutputDirectory 'stale-totals.csv'
$rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
$staleRegisters = @($rows | Where-Object stale_rows -gt 0 | Select-Object -ExpandProperty register -Unique)
$summary = [ordered]@{
    generated_at_local=(Get-Date).ToString('s'); server=$Server; database=$Database
    organization_table=$organizationTable; tables_checked=@($rows.physical_table | Select-Object -Unique).Count
    stale_rows=[long]($rows | Measure-Object stale_rows -Sum).Sum; stale_registers=$staleRegisters; csv=$path
}
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 4
