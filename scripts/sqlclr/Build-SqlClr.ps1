[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'out')
)

$ErrorActionPreference = 'Stop'

$compilerCandidates = @(
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
)
$compiler = $compilerCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $compiler) {
    throw 'The .NET Framework C# compiler was not found.'
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$assemblyPath = Join-Path $OutputDirectory 'OneC.SqlClr.dll'
$sourcePath = Join-Path $PSScriptRoot 'OneC.SqlClr.cs'
$templatePath = Join-Path $PSScriptRoot 'install-template.sql'
$installPath = Join-Path $OutputDirectory 'install-OneC.SqlClr.sql'

& $compiler /nologo /target:library /optimize+ /platform:anycpu `
    /reference:System.dll /reference:System.Data.dll `
    /out:$assemblyPath $sourcePath
if ($LASTEXITCODE -ne 0) {
    throw "csc.exe failed with exit code $LASTEXITCODE"
}

$assemblyHex = '0x' + ([BitConverter]::ToString([IO.File]::ReadAllBytes($assemblyPath)) -replace '-', '')
$template = [IO.File]::ReadAllText($templatePath, [Text.Encoding]::UTF8)
if (-not $template.Contains('__ASSEMBLY_HEX__')) {
    throw 'Assembly placeholder was not found in install-template.sql.'
}

[IO.File]::WriteAllText(
    $installPath,
    $template.Replace('__ASSEMBLY_HEX__', $assemblyHex),
    [Text.UTF8Encoding]::new($false))

[pscustomobject]@{
    Assembly = $assemblyPath
    InstallScript = $installPath
    AssemblyBytes = (Get-Item -LiteralPath $assemblyPath).Length
    Sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $assemblyPath).Hash
}
