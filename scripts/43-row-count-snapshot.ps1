[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$Database,
    [Parameter(Mandatory)] [string]$Label,
    [string]$CatalogDirectory,
    [string]$CompareWith,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-row-snapshot'))
)
<#
    Read-only row-count snapshot of every user table, for the
    "before delete / after delete / after 1C update" comparison.

    Counts come from sys.dm_db_partition_stats (heap or clustered index only),
    which is exact for committed rows and takes seconds on the whole base.
    With -CompareWith <previous row-counts.csv> the script also writes
    row-count-diff.csv: tables whose count changed, with the 1C logical name
    from the catalog of script 40.  Tables that grow again after a 1C update
    are what the platform recreates; they are candidates for PROTECT.
#>
$ErrorActionPreference = 'Stop'

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

$typeNames = @{}
if ($CatalogDirectory) {
    foreach ($row in Import-Csv -LiteralPath (Join-Path $CatalogDirectory 'type-names.csv')) { $typeNames[$row.type_table] = $row.logical_name }
}
function Get-LogicalName([string]$Table) {
    $name = $Table.TrimStart('_') -replace 'X[0-9]+$',''
    if ($name -match '^(?<owner>[A-Za-z]+[0-9]+)_(?<part>[A-Za-z]+[0-9]+)') {
        return (@($typeNames[$Matches.owner], $typeNames[$Matches.part]) | Where-Object { $_ }) -join '.'
    }
    return $typeNames[$name]
}

$connection = [System.Data.SqlClient.SqlConnection]::new(
    "Server=$Server;Database=$Database;Integrated Security=True;Encrypt=False;Application Name=Codex-RowSnapshot")
try {
    $connection.Open()
    $command = $connection.CreateCommand()
    $command.CommandTimeout = 600
    $command.CommandText = @'
SET NOCOUNT ON;
SELECT t.name AS table_name, SUM(p.row_count) AS rows_
FROM sys.tables t
JOIN sys.dm_db_partition_stats p ON p.object_id=t.object_id AND p.index_id IN (0,1)
WHERE t.is_ms_shipped=0
GROUP BY t.name
ORDER BY t.name;
'@
    $adapter = [System.Data.SqlClient.SqlDataAdapter]::new($command)
    $table = [System.Data.DataTable]::new()
    [void]$adapter.Fill($table)
    $adapter.Dispose(); $command.Dispose()
}
finally { $connection.Dispose() }

$snapshot = @($table.Rows | ForEach-Object { [pscustomobject]@{ table=[string]$_.table_name; rows=[int64]$_.rows_; logical_name=Get-LogicalName ([string]$_.table_name) } })
$snapshotPath = Join-Path $OutputDirectory 'row-counts.csv'
$snapshot | Export-Csv -LiteralPath $snapshotPath -NoTypeInformation -Encoding UTF8

$summary = [ordered]@{
    generated_at_local=(Get-Date).ToString('s'); server=$Server; database=$Database; label=$Label; read_only=$true
    tables=$snapshot.Count; total_rows=[int64](($snapshot | Measure-Object rows -Sum).Sum)
    non_empty_tables=@($snapshot | Where-Object rows -gt 0).Count
    row_counts=$snapshotPath; row_counts_sha256=(Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash.ToLowerInvariant()
}
if ($CompareWith) {
    $previous = @{}
    foreach ($row in Import-Csv -LiteralPath $CompareWith) { $previous[$row.table] = [int64]$row.rows }
    $diff = foreach ($row in $snapshot) {
        $before = if ($previous.ContainsKey($row.table)) { $previous[$row.table] } else { $null }
        if ($null -eq $before -or $before -ne $row.rows) {
            [pscustomobject]@{ table=$row.table; logical_name=$row.logical_name; before=$before; after=$row.rows; delta=$row.rows - [int64]$before }
        }
    }
    foreach ($name in $previous.Keys | Where-Object { $_ -notin $snapshot.table }) {
        $diff += [pscustomobject]@{ table=$name; logical_name=Get-LogicalName $name; before=$previous[$name]; after=$null; delta=-$previous[$name] }
    }
    $diffPath = Join-Path $OutputDirectory 'row-count-diff.csv'
    @($diff | Sort-Object { [Math]::Abs([int64]$_.delta) } -Descending) | Export-Csv -LiteralPath $diffPath -NoTypeInformation -Encoding UTF8
    $summary.compared_with = (Resolve-Path -LiteralPath $CompareWith).Path
    $summary.changed_tables = @($diff).Count
    $summary.grown_tables = @($diff | Where-Object { [int64]$_.delta -gt 0 }).Count
    $summary.shrunk_tables = @($diff | Where-Object { [int64]$_.delta -lt 0 }).Count
    $summary.row_delta = [int64](($diff | Measure-Object delta -Sum).Sum)
    $summary.row_count_diff = $diffPath
}
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 4
