[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [Parameter(Mandatory)] [string]$Database,
    [switch]$Shrink,
    [ValidateRange(0,90)] [int]$ShrinkFreePercent = 10,
    [switch]$SetSimpleRecovery,
    [switch]$RebuildIndexes,
    [ValidateRange(1,100)] [int]$RebuildFragmentationPercent = 30,
    [ValidateRange(1,10000000)] [int]$MinPages = 1000,
    [switch]$UpdateStatistics,
    [string]$ConfirmationToken,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-sql-maintenance'))
)
<#
    MS SQL maintenance of a 1C database after a mass deletion.

    Without switches the script only reports: file sizes and free space, and
    fragmentation of indexes with at least -MinPages pages (LIMITED scan).

      -Shrink            DBCC SHRINKFILE of the data file down to used space plus
                         -ShrinkFreePercent, and of the log file.  For a copy that
                         is handed over; requires -ConfirmationToken SHRINK-FILES.
      -SetSimpleRecovery with -Shrink: switch the database to SIMPLE recovery
                         first.  In FULL recovery without log backups the log
                         cannot be truncated, and shrinking the data file makes
                         it grow.  Breaks the log backup chain, so only for a copy.
      -RebuildIndexes    ALTER INDEX ... REBUILD for indexes fragmented above
                         -RebuildFragmentationPercent (run after -Shrink, which
                         fragments indexes).
      -UpdateStatistics  sp_updatestats.

    Order when combined: shrink, rebuild, statistics.  1C needs no exclusive
    mode for this, but users should not work in the base meanwhile.
#>
$ErrorActionPreference = 'Stop'
if ($Shrink -and $ConfirmationToken -ne 'SHRINK-FILES') { throw '-Shrink requires -ConfirmationToken SHRINK-FILES.' }
if ($SetSimpleRecovery -and -not $Shrink) { throw '-SetSimpleRecovery works only with -Shrink.' }

function Invoke-Table([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [int]$Timeout = 0) {
    $command = $Connection.CreateCommand()
    $command.CommandTimeout = $Timeout
    $command.CommandText = "SET NOCOUNT ON;`n$Sql"
    try {
        $adapter = [System.Data.SqlClient.SqlDataAdapter]::new($command)
        $table = [System.Data.DataTable]::new()
        [void]$adapter.Fill($table)
        $adapter.Dispose()
        Write-Output -NoEnumerate $table
    }
    finally { $command.Dispose() }
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$connection = [System.Data.SqlClient.SqlConnection]::new("Server=$Server;Database=$Database;Integrated Security=True;Encrypt=False;Application Name=Codex-SqlMaintenance")
$filesSql = @'
SELECT name, type_desc, physical_name, size/128 AS size_mb,
       FILEPROPERTY(name,'SpaceUsed')/128 AS used_mb, (size-FILEPROPERTY(name,'SpaceUsed'))/128 AS free_mb
FROM sys.database_files;
'@
$fragmentationSql = @"
SELECT OBJECT_NAME(s.object_id) AS table_name, i.name AS index_name, s.index_id,
       CONVERT(decimal(5,1), s.avg_fragmentation_in_percent) AS fragmentation, s.page_count
FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'LIMITED') s
JOIN sys.indexes i ON i.object_id=s.object_id AND i.index_id=s.index_id
WHERE s.index_id>0 AND s.page_count>=$MinPages AND s.alloc_unit_type_desc=N'IN_ROW_DATA'
ORDER BY s.avg_fragmentation_in_percent*s.page_count DESC;
"@
$started = Get-Date
try {
    $connection.Open()
    if ([string](Invoke-Table $connection 'SELECT DB_NAME() AS n;').Rows[0].n -ne $Database) { throw "Connected to an unexpected database." }
    $filesBefore = Invoke-Table $connection $filesSql
    $fragmentation = Invoke-Table $connection $fragmentationSql
    $filesBefore.Rows | Select-Object name,type_desc,physical_name,size_mb,used_mb,free_mb | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'files-before.csv') -NoTypeInformation -Encoding UTF8
    $fragmentation.Rows | Select-Object table_name,index_name,index_id,fragmentation,page_count | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'fragmentation-before.csv') -NoTypeInformation -Encoding UTF8
    $toRebuild = @($fragmentation.Rows | Where-Object { [decimal]$_.fragmentation -ge $RebuildFragmentationPercent })
    $actions = [System.Collections.Generic.List[object]]::new()

    $recoveryBefore = [string](Invoke-Table $connection 'SELECT recovery_model_desc AS m FROM sys.databases WHERE database_id=DB_ID();').Rows[0].m
    if ($Shrink) {
        if ($SetSimpleRecovery -and $recoveryBefore -ne 'SIMPLE') {
            [void](Invoke-Table $connection "ALTER DATABASE [$($Database.Replace(']',']]'))] SET RECOVERY SIMPLE;")
            $actions.Add([pscustomobject]@{ action='recovery'; target="$recoveryBefore -> SIMPLE"; seconds=0 })
            Write-Host "Recovery $recoveryBefore -> SIMPLE"
        }
        foreach ($file in $filesBefore.Rows) {
            $target = if ($file.type_desc -eq 'ROWS') { [int]([double]$file.used_mb * (1 + $ShrinkFreePercent / 100.0)) } else { 1024 }
            $t0 = Get-Date
            # Shrinking the data file fills the log; in SIMPLE recovery a checkpoint frees it.
            if ($SetSimpleRecovery -and $file.type_desc -eq 'LOG') { [void](Invoke-Table $connection 'CHECKPOINT;') }
            [void](Invoke-Table $connection "DBCC SHRINKFILE (N'$($file.name)', $target) WITH NO_INFOMSGS;")
            $actions.Add([pscustomobject]@{ action='shrink'; target="$($file.name) -> $target MB"; seconds=[Math]::Round(((Get-Date)-$t0).TotalSeconds,1) })
            Write-Host ("Shrink {0} to {1} MB" -f $file.name, $target)
        }
        # Shrinking moves pages and fragments indexes: measure again.
        $toRebuild = @((Invoke-Table $connection $fragmentationSql).Rows | Where-Object { [decimal]$_.fragmentation -ge $RebuildFragmentationPercent })
    }
    if ($RebuildIndexes) {
        $i = 0
        foreach ($index in $toRebuild) {
            $i++
            if ([string]$index.table_name -notmatch '^[A-Za-z0-9_]+$' -or [string]$index.index_name -notmatch '^[A-Za-z0-9_]+$') { continue }
            $t0 = Get-Date
            [void](Invoke-Table $connection "ALTER INDEX [$($index.index_name)] ON dbo.[$($index.table_name)] REBUILD;")
            $actions.Add([pscustomobject]@{ action='rebuild'; target="$($index.table_name).$($index.index_name) ($($index.fragmentation)%)"; seconds=[Math]::Round(((Get-Date)-$t0).TotalSeconds,1) })
            if (($i % 25) -eq 0) { Write-Host ("Rebuild {0} of {1}" -f $i, $toRebuild.Count) }
        }
    }
    if ($UpdateStatistics) {
        $t0 = Get-Date
        [void](Invoke-Table $connection 'EXEC sp_updatestats;')
        $actions.Add([pscustomobject]@{ action='statistics'; target='sp_updatestats'; seconds=[Math]::Round(((Get-Date)-$t0).TotalSeconds,1) })
    }
    $filesAfter = Invoke-Table $connection $filesSql
    $filesAfter.Rows | Select-Object name,type_desc,physical_name,size_mb,used_mb,free_mb | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'files-after.csv') -NoTypeInformation -Encoding UTF8
    $actions | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'actions.csv') -NoTypeInformation -Encoding UTF8

    $summary = [ordered]@{
        generated_at_local=(Get-Date).ToString('s'); server=$Server; database=$Database
        changed=[bool]($Shrink -or $RebuildIndexes -or $UpdateStatistics)
        recovery_before=$recoveryBefore
        recovery_after=[string](Invoke-Table $connection 'SELECT recovery_model_desc AS m FROM sys.databases WHERE database_id=DB_ID();').Rows[0].m
        files_before=@($filesBefore.Rows | ForEach-Object { [ordered]@{ name=$_.name; size_mb=$_.size_mb; used_mb=$_.used_mb; free_mb=$_.free_mb } })
        files_after=@($filesAfter.Rows | ForEach-Object { [ordered]@{ name=$_.name; size_mb=$_.size_mb; used_mb=$_.used_mb; free_mb=$_.free_mb } })
        large_indexes=$fragmentation.Rows.Count
        fragmented_indexes=@($fragmentation.Rows | Where-Object { [decimal]$_.fragmentation -ge $RebuildFragmentationPercent }).Count
        actions=$actions.Count
        duration_seconds=[Math]::Round(((Get-Date)-$started).TotalSeconds,1)
    }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
    $summary | ConvertTo-Json -Depth 5
}
finally { $connection.Dispose() }
