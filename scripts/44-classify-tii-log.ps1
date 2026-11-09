[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$TiiLog,
    [Parameter(Mandatory)] [string]$DeleteManifest,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-tii-classification'))
)
<#
    Classifies the log of 1C "Тестирование и исправление" run in test-only mode
    (DESIGNER /IBCheckAndRepair -TestOnly ... /Out <log>) against the frozen
    delete manifest of a run.

    Every finding "Проверка логической целостности. <Объект>.<Вид>.<Поле> ..."
    is grouped by object path, field and message.  A reference to a missing
    object, printed as (<type code>:<32 hex>), is split into
      deleted-by-run   the target is in the delete manifest: a miss of the
                       KEEP/DELETE boundary that must be fixed in scripts or rules;
      pre-existing     the target is not in the manifest: the reference was
                       already broken before the run (source data defect).
#>
$ErrorActionPreference = 'Stop'

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

$deleted = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($row in Import-Csv -LiteralPath $DeleteManifest) {
    [void]$deleted.Add([string]$row.type_code + ':' + ([string]$row.id_hex).Substring(2))
}

$lines = Get-Content -LiteralPath $TiiLog -Encoding UTF8
$findings = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -notmatch '^Проверка (?<check>[^.]+)\. (?<path>[^ ]+)(?<rest>.*)$') { continue }
    $check = $Matches.check
    $path = $Matches.path
    $rest = $Matches.rest
    # The message is the first following line that is not a common-attribute line.
    $message = ''
    for ($j = $i + 1; $j -lt [Math]::Min($i + 6, $lines.Count); $j++) {
        if ($lines[$j] -match '^ОбщийРеквизит\.' -or [string]::IsNullOrWhiteSpace($lines[$j])) { continue }
        if ($lines[$j] -match '^Проверка ') { break }
        $message = $lines[$j].Trim()
        break
    }
    $refs = @([regex]::Matches($rest, '\((?<code>[0-9]+):(?<id>[0-9a-f]{32})\)') | ForEach-Object { "$($_.Groups['code'].Value):$($_.Groups['id'].Value)" })
    $origin = if ($refs.Count -eq 0) { 'no-reference' }
              elseif (@($refs | Where-Object { $deleted.Contains($_) }).Count -gt 0) { 'deleted-by-run' }
              else { 'pre-existing' }
    $findings.Add([pscustomobject]@{
        check=$check; object=($path -replace '\.[^.]+\.[^.]+$',''); field=$path; message=$message; origin=$origin
        reference=($refs -join ';'); presentation=$rest.Trim()
    })
}

$findingsPath = Join-Path $OutputDirectory 'tii-findings.csv'
$findings | Export-Csv -LiteralPath $findingsPath -NoTypeInformation -Encoding UTF8
$groupsPath = Join-Path $OutputDirectory 'tii-groups.csv'
@($findings | Group-Object field, message, origin | ForEach-Object {
    $first = $_.Group[0]
    [pscustomobject]@{ findings=$_.Count; origin=$first.origin; field=$first.field; message=$first.message
                       distinct_references=@($_.Group.reference | Where-Object { $_ } | Sort-Object -Unique).Count; example=$first.presentation }
} | Sort-Object findings -Descending) | Export-Csv -LiteralPath $groupsPath -NoTypeInformation -Encoding UTF8

$summary = [ordered]@{
    generated_at_local=(Get-Date).ToString('s'); tii_log=(Resolve-Path -LiteralPath $TiiLog).Path
    delete_manifest=(Resolve-Path -LiteralPath $DeleteManifest).Path
    findings=$findings.Count
    by_origin=[ordered]@{}
    deleted_by_run_fields=@($findings | Where-Object origin -eq 'deleted-by-run' | Group-Object field | ForEach-Object { [ordered]@{ field=$_.Name; findings=$_.Count } })
    findings_csv=$findingsPath; groups_csv=$groupsPath
}
foreach ($g in @($findings | Group-Object origin)) { $summary.by_origin[$g.Name] = $g.Count }
$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 5
