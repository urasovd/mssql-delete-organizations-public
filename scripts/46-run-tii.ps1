[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$InfoBase,
    [string]$User = 'admin',
    [string]$Password = '',
    [ValidateSet('Check','Totals','CheckAndTotals','Repair','Rebuild')]
    [string]$Profile = 'Check',
    [string]$PlatformDirectory = 'C:\Program Files\1cv8\8.3.27.2214\bin',
    [string]$Database,
    [string]$SqlServer = '.',
    [string]$CatalogDirectory,
    [string]$DeleteManifest,
    [string]$TimeLimit = '004:00',
    [switch]$UseStartPoint,
    [switch]$PrepareCluster,
    [string]$ConfirmationToken,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-tii'))
)
<#
    Runs 1C "Тестирование и исправление" in batch Designer mode
    (DESIGNER /IBCheckAndRepair) with a named set of options, waits for it and
    classifies the log.

    InfoBase is the server infobase path, e.g. server\copy.  The
    infobase must be free of other sessions: TiI needs exclusive mode.

    Profiles (the "галочки" of the dialog):
      Check          logical + referential integrity, test only, nothing changes;
                     the log is classified by script 44 against -DeleteManifest
                     (findings caused by the deletion vs pre-existing ones).
      Totals         recalculation of totals only.
      CheckAndTotals Check, then Totals, as two runs.
      Repair         logical + referential integrity WITH fixes: references to
                     missing objects are cleared, objects with partly lost data
                     are deleted.  Changes data; requires -ConfirmationToken
                     TII-REPAIR and is meant for a copy whose Check log was
                     reviewed.
      Rebuild        reindex + rebuild (restructuring) of all tables.  Long; on
                     MS SQL usually replaced by script 48.

    -PrepareCluster (rac of the same platform): before every run deny scheduled
    jobs of the infobase and terminate all its sessions.  Without it scheduled
    jobs started by the cluster during TiI ended the Designer with exit -1
    on bit_copy.

    Table compression (-IBCompression) applies to file infobases only and is not
    offered.  With -Database and -CatalogDirectory a row-count snapshot is
    taken before and after a changing profile (script 43).
#>
$ErrorActionPreference = 'Stop'

$designer = Join-Path $PlatformDirectory '1cv8.exe'
if (-not (Test-Path -LiteralPath $designer -PathType Leaf)) { throw "1C Designer was not found: $designer" }
if ($Profile -eq 'Repair' -and $ConfirmationToken -ne 'TII-REPAIR') { throw 'Profile Repair changes data and requires -ConfirmationToken TII-REPAIR.' }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

# @(...) keeps a one-run profile as a list of one [name, options] pair.
$runs = @(switch ($Profile) {
    'Check' { ,@('Check', @('-LogAndRefsIntegrity','-TestOnly','-BadRefNone','-BadDataNone')) }
    'Totals' { ,@('Totals', @('-RecalcTotals')) }
    'CheckAndTotals' { @(@('Check', @('-LogAndRefsIntegrity','-TestOnly','-BadRefNone','-BadDataNone')), @('Totals', @('-RecalcTotals'))) }
    'Repair' { ,@('Repair', @('-LogAndRefsIntegrity','-BadRefClear','-BadDataDelete')) }
    'Rebuild' { ,@('Rebuild', @('-ReIndex','-Rebuild')) }
})

function Invoke-Snapshot([string]$Label) {
    if (-not ($Database -and $CatalogDirectory)) { return $null }
    $dir = Join-Path $OutputDirectory "snapshot-$Label"
    & (Join-Path $PSScriptRoot '43-row-count-snapshot.ps1') -Server $SqlServer -Database $Database -Label $Label `
        -CatalogDirectory $CatalogDirectory -OutputDirectory $dir | Out-Null
    return (Join-Path $dir 'row-counts.csv')
}

function Invoke-PrepareCluster {
    $rac = Join-Path $PlatformDirectory 'rac.exe'
    if ($InfoBase -notmatch '^(?<host>[^\\:]+)(:\d+)?\\(?<name>.+)$') { throw "InfoBase must be <server>\<name>: $InfoBase" }
    $ibName = $Matches.name
    $cluster = ((& $rac cluster list) | Select-String '^cluster\s*:\s*(\S+)' | Select-Object -First 1).Matches[0].Groups[1].Value
    $lines = & $rac infobase summary list --cluster=$cluster
    $ibId = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^infobase\s*:\s*(\S+)') { $candidate = $Matches[1] }
        if ($lines[$i] -match '^name\s*:\s*"?(?<n>[^"]+)"?\s*$' -and $Matches.n -eq $ibName) { $ibId = $candidate; break }
    }
    if (-not $ibId) { throw "Infobase $ibName is not registered in cluster $cluster." }
    & $rac infobase update --cluster=$cluster --infobase=$ibId --infobase-user=$User --infobase-pwd=$Password --scheduled-jobs-deny=on | Out-Null
    $sessions = @((& $rac session list --cluster=$cluster --infobase=$ibId) | Select-String '^session\s*:\s*(\S+)' | ForEach-Object { $_.Matches[0].Groups[1].Value })
    foreach ($session in $sessions) { & $rac session terminate --cluster=$cluster --session=$session | Out-Null }
    Write-Host ("Cluster: scheduled jobs denied, {0} sessions terminated" -f $sessions.Count)
    return $sessions.Count
}

$results = [System.Collections.Generic.List[object]]::new()
foreach ($run in $runs) {
    $terminated = if ($PrepareCluster) { Invoke-PrepareCluster } else { $null }
    $name = [string]$run[0]
    $options = @($run[1])
    if ($UseStartPoint) { $options += '-UseStartPoint' }
    if ($TimeLimit) { $options += "-TimeLimit:$TimeLimit" }
    $log = Join-Path $OutputDirectory "tii-$($name.ToLowerInvariant()).log"
    $changing = $name -ne 'Check'
    $before = if ($changing) { Invoke-Snapshot "before-$($name.ToLowerInvariant())" } else { $null }

    $arguments = @('DESIGNER', '/S', $InfoBase, '/N', $User, '/P', $Password, '/DisableStartupDialogs', '/Out', $log, '/IBCheckAndRepair') + $options
    $quoted = $arguments | ForEach-Object { if ($_ -eq '' -or $_ -match '[\s"]') { '"' + $_.Replace('"','\"') + '"' } else { $_ } }
    Write-Host ("TiI {0}: {1}" -f $name, ($options -join ' '))
    $started = Get-Date
    $process = Start-Process -FilePath $designer -ArgumentList ($quoted -join ' ') -PassThru -Wait
    $seconds = [Math]::Round(((Get-Date) - $started).TotalSeconds, 1)
    $text = if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log -Encoding UTF8 } else { @() }
    $finished = [bool]($text | Where-Object { $_ -match '^Тестирование закончено' })
    $exclusive = [bool]($text | Where-Object { $_ -match 'монопольн' })
    # A server on a developer license allows very few sessions: another client
    # (a Designer of a different base included) makes TiI fail at start or later.
    $license = [bool]($text | Where-Object { $_ -match 'лицензи' })

    $result = [ordered]@{ run=$name; options=($options -join ' '); sessions_terminated=$terminated; exit_code=$process.ExitCode; seconds=$seconds
                          finished=$finished; exclusive_mode_error=$exclusive; license_error=$license; log=$log; log_lines=$text.Count }
    if ($name -eq 'Check' -and $DeleteManifest -and $finished) {
        $classification = & (Join-Path $PSScriptRoot '44-classify-tii-log.ps1') -TiiLog $log -DeleteManifest $DeleteManifest `
            -OutputDirectory (Join-Path $OutputDirectory 'classification') | ConvertFrom-Json
        $result.findings = $classification.findings
        $result.by_origin = $classification.by_origin
        $result.deleted_by_run_fields = $classification.deleted_by_run_fields
    }
    if ($changing -and $before) {
        $after = Invoke-Snapshot "after-$($name.ToLowerInvariant())"
        $diffDir = Join-Path $OutputDirectory "diff-$($name.ToLowerInvariant())"
        $diff = & (Join-Path $PSScriptRoot '43-row-count-snapshot.ps1') -Server $SqlServer -Database $Database -Label "diff-$name" `
            -CatalogDirectory $CatalogDirectory -CompareWith $before -OutputDirectory $diffDir | ConvertFrom-Json
        $result.changed_tables = $diff.changed_tables
        $result.row_delta = $diff.row_delta
        $result.row_count_diff = $diff.row_count_diff
    }
    $results.Add([pscustomobject]$result)
    Write-Host ("TiI {0}: finished={1}, {2} s, exit {3}" -f $name, $finished, $seconds, $process.ExitCode)
    if (-not $finished) { break }
}

$summary = [ordered]@{
    generated_at_local=(Get-Date).ToString('s'); infobase=$InfoBase; user=$User; profile=$Profile
    runs=@($results)
    result=if (@($results | Where-Object { -not $_.finished }).Count -eq 0 -and $results.Count -eq @($runs).Count) { 'FINISHED' } else { 'INCOMPLETE' }
}
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 6
