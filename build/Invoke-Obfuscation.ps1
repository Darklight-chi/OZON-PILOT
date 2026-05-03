param(
    [Parameter(Mandatory = $true)]
    [string]$ObfuscarExe,

    [Parameter(Mandatory = $true)]
    [string]$InputDirectory,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $true)]
    [string[]]$Modules
)

$ErrorActionPreference = "Stop"

function Escape-Xml([string]$value) {
    return [System.Security.SecurityElement]::Escape($value)
}

$toolValue = $ObfuscarExe
if ($null -eq $toolValue) {
    $toolValue = ""
}

$inputValue = $InputDirectory
if ($null -eq $inputValue) {
    $inputValue = ""
}

$outputValue = $OutputDirectory
if ($null -eq $outputValue) {
    $outputValue = ""
}

$toolPath = [System.IO.Path]::GetFullPath($toolValue.Trim().Trim('"'))
$inputPath = [System.IO.Path]::GetFullPath($inputValue.Trim().Trim('"'))
$outputPath = [System.IO.Path]::GetFullPath($outputValue.Trim().Trim('"'))

if (!(Test-Path $toolPath)) {
    throw "Obfuscar executable not found: $toolPath"
}

if (!(Test-Path $inputPath)) {
    throw "Input directory not found: $inputPath"
}

if (Test-Path $outputPath) {
    Remove-Item -LiteralPath $outputPath -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $outputPath | Out-Null

$moduleLines = foreach ($module in $Modules) {
    $modulePath = [System.IO.Path]::GetFullPath((Join-Path $inputPath $module))
    if (!(Test-Path $modulePath)) {
        throw "Module not found: $modulePath"
    }

    '  <Module file="{0}" />' -f (Escape-Xml $modulePath)
}

$configPath = Join-Path $outputPath "obfuscar.xml"
$configContent = @(
    '<?xml version="1.0" encoding="utf-8"?>'
    '<Obfuscator>'
    ('  <Var name="InPath" value="{0}" />' -f (Escape-Xml $inputPath))
    ('  <Var name="OutPath" value="{0}" />' -f (Escape-Xml $outputPath))
    '  <Var name="KeepPublicApi" value="true" />'
    '  <Var name="HidePrivateApi" value="true" />'
    '  <Var name="ReuseNames" value="false" />'
    '  <Var name="UseUnicodeNames" value="false" />'
    '  <Var name="SkipGenerated" value="true" />'
    $moduleLines
    '</Obfuscator>'
) -join [Environment]::NewLine

Set-Content -LiteralPath $configPath -Value $configContent -Encoding UTF8

Push-Location (Split-Path -Parent $toolPath)
try {
    & $toolPath $configPath
    if ($LASTEXITCODE -ne 0) {
        throw "Obfuscar failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}
