[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$InfoBase,
    [string]$TargetInfoBase,
    [string]$User = 'admin',
    [string]$Password = '',
    [Parameter(Mandatory)] [string]$Database,
    [string]$TargetDatabase,
    [string]$SqlServer = 'localhost',
    [Parameter(Mandatory)] [string]$CatalogDirectory,
    [Parameter(Mandatory)] [string]$DtPath,
    [string]$PlatformDirectory = 'C:\Program Files\1cv8\8.3.27.2214\bin',
    [string]$ConfirmationToken,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-reload-dt'))
)
<#
    Dump the infobase to .dt and load it back (into the same infobase or into
    -TargetInfoBase), then compare row counts table by table.

    Deleting in SQL leaves the platform table BinaryData with the binary data
    of deleted objects (attached files, object versions, value storages): only
    compressed values of other tables point to its keys.  A .dt holds only the
    data that is referenced, so the reload drops the rest (BIT: 85 GB -> 6 GB).

    The comparison must show changes only in BinaryData and the service table
    _ExtensionsRestructNGS; anything else stops the script.  Loading into the
    same infobase replaces its data: take a backup first (Apply's backup is
    of the copy before deletion).  The file of the database keeps its size:
    shrink it with 48 afterwards.

    A new target infobase must be created with the same date offset as the
    source (rac infobase create ... --date-offset=2000 for most 1C bases on
    MS SQL), otherwise dates differ in SQL and 45 cannot compare the bases.

    Requires exclusive access: sessions of both infobases are terminated and
    scheduled jobs denied (rac).  Token RELOAD-DT.
#>
$ErrorActionPreference = 'Stop'
if ($ConfirmationToken -ne 'RELOAD-DT') { throw 'Reloading replaces the data of the target infobase: pass -ConfirmationToken RELOAD-DT.' }
if (-not $TargetInfoBase) { $TargetInfoBase = $InfoBase }
if (-not $TargetDatabase) { $TargetDatabase = $Database }
$designer = Join-Path $PlatformDirectory '1cv8.exe'
$rac = Join-Path $PlatformDirectory 'rac.exe'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

function Close-Sessions([string]$Path) {
    if ($Path -notmatch '^(?<host>[^\\:]+)(:\d+)?\\(?<name>.+)$') { throw "InfoBase must be <server>\<name>: $Path" }
    $name = $Matches.name
    $cluster = ((& $rac cluster list) | Select-String '^cluster\s*:\s*(\S+)' | Select-Object -First 1).Matches[0].Groups[1].Value
    $lines = & $rac infobase summary list --cluster=$cluster
    $id = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^infobase\s*:\s*(\S+)') { $candidate = $Matches[1] }
        if ($lines[$i] -match '^name\s*:\s*"?(?<n>[^"]+)"?\s*$' -and $Matches.n -eq $name) { $id = $candidate; break }
    }
    if (-not $id) { throw "Infobase $name is not registered in cluster $cluster." }
    & $rac infobase update --cluster=$cluster --infobase=$id --infobase-user=$User --infobase-pwd=$Password --scheduled-jobs-deny=on | Out-Null
    $sessions = @((& $rac session list --cluster=$cluster --infobase=$id) | Select-String '^session\s*:\s*(\S+)' | ForEach-Object { $_.Matches[0].Groups[1].Value })
    foreach ($session in $sessions) { & $rac session terminate --cluster=$cluster --session=$session | Out-Null }
    return $sessions.Count
}

function Invoke-Designer([string]$Path, [string[]]$Arguments, [string]$Log) {
    $credentials = if ($User) { @('/N', $User, '/P', "`"$Password`"") } else { @() }
    $process = Start-Process -FilePath $designer -Wait -PassThru -ArgumentList (@('DESIGNER', '/S', $Path) + $credentials +
        @('/DisableStartupDialogs', '/DisableStartupMessages') + $Arguments + @('/Out', "`"$Log`""))
    $text = if (Test-Path -LiteralPath $Log) { (Get-Content -LiteralPath $Log -Raw).Trim() } else { '' }
    if ($process.ExitCode -ne 0) { throw "Designer exit $($process.ExitCode): $text" }
    return $text
}

$summary = [ordered]@{ generated_at_local = (Get-Date).ToString('s'); infobase = $InfoBase; target_infobase = $TargetInfoBase; dt = $DtPath }
$watch = [Diagnostics.Stopwatch]::StartNew()

& (Join-Path $PSScriptRoot '43-row-count-snapshot.ps1') -Server $SqlServer -Database $Database -Label before-dt `
    -CatalogDirectory $CatalogDirectory -OutputDirectory (Join-Path $OutputDirectory 'before') | Out-Null

$summary.sessions_terminated = Close-Sessions $InfoBase
$summary.dump = Invoke-Designer $InfoBase @('/DumpIB', "`"$DtPath`"") (Join-Path $OutputDirectory 'dump.log')
$summary.dt_bytes = (Get-Item -LiteralPath $DtPath).Length
$summary.dump_minutes = [Math]::Round($watch.Elapsed.TotalMinutes, 1)

if ($TargetInfoBase -ne $InfoBase) { [void](Close-Sessions $TargetInfoBase) } else { [void](Close-Sessions $InfoBase) }
$summary.restore = Invoke-Designer $TargetInfoBase @('/RestoreIB', "`"$DtPath`"") (Join-Path $OutputDirectory 'restore.log')
$summary.total_minutes = [Math]::Round($watch.Elapsed.TotalMinutes, 1)

& (Join-Path $PSScriptRoot '43-row-count-snapshot.ps1') -Server $SqlServer -Database $TargetDatabase -Label after-dt `
    -CatalogDirectory $CatalogDirectory -CompareWith (Join-Path $OutputDirectory 'before\row-counts.csv') `
    -OutputDirectory (Join-Path $OutputDirectory 'after') | Out-Null
$diff = @(Import-Csv -LiteralPath (Join-Path $OutputDirectory 'after\row-count-diff.csv'))
$allowed = 'BinaryData', '_ExtensionsRestructNGS'
# The restore may create empty platform tables (BinaryDataStorage*, ExternalBinDataStrgs*): 0 rows is no change
$unexpected = @($diff | Where-Object { $allowed -notcontains $_.table -and [int64]('0' + $_.before) -ne [int64]('0' + $_.after) })
$binary = $diff | Where-Object table -eq 'BinaryData'
$summary.binary_data_rows_before = if ($binary) { [int64]$binary.before } else { $null }
$summary.binary_data_rows_after = if ($binary) { [int64]$binary.after } else { $null }
$summary.unexpected_changes = $unexpected.Count
$summary.result = if ($unexpected.Count) { 'FAIL' } else { 'PASS' }

$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 4
if ($unexpected.Count) { throw "Row counts changed outside BinaryData in $($unexpected.Count) table(s); see $OutputDirectory\after\row-count-diff.csv." }
