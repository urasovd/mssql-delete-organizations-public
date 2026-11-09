[CmdletBinding()]
param(
    [string]$PlatformDirectory = 'C:\Program Files\1cv8\8.3.27.2214\bin',
    [string]$SourceDirectory = (Join-Path $PSScriptRoot '..\1c'),
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\build\1c')
)
<#
    Builds the external data processors in 1c/ (XML sources) into .epf files.

    Uses a throw-away file infobase, so no server infobase or 1C user is
    needed.  The processors are not compiled against a configuration here:
    their code is compiled when they are opened in the target infobase.
#>
$ErrorActionPreference = 'Stop'

$designer = Join-Path $PlatformDirectory '1cv8.exe'
if (-not (Test-Path -LiteralPath $designer -PathType Leaf)) { throw "1C Designer was not found: $designer" }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

$base = Join-Path ([IO.Path]::GetTempPath()) ('epf-build-' + [guid]::NewGuid().ToString('N'))
$log = Join-Path $OutputDirectory 'build.log'
try {
    $process = Start-Process -FilePath $designer -ArgumentList @('CREATEINFOBASE', "File=`"$base`"", '/Out', "`"$log`"", '/DisableStartupDialogs') -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "Temporary infobase was not created (exit $($process.ExitCode)): $(Get-Content -LiteralPath $log -Raw)" }

    foreach ($root in Get-ChildItem -LiteralPath $SourceDirectory -Filter '*.xml' -File) {
        $epf = Join-Path $OutputDirectory ($root.BaseName + '.epf')
        $process = Start-Process -FilePath $designer -Wait -PassThru -ArgumentList @(
            'DESIGNER', '/F', "`"$base`"", '/DisableStartupDialogs', '/DisableStartupMessages',
            '/LoadExternalDataProcessorOrReportFromFiles', "`"$($root.FullName)`"", "`"$epf`"", '/Out', "`"$log`"")
        $text = if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log -Raw } else { '' }
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $epf)) { throw "$($root.Name) was not built (exit $($process.ExitCode)): $text" }
        [pscustomobject]@{ epf = $epf; bytes = (Get-Item -LiteralPath $epf).Length; log = $text.Trim() }
    }
} finally {
    if (Test-Path -LiteralPath $base) { [IO.Directory]::Delete($base, $true) }
}
