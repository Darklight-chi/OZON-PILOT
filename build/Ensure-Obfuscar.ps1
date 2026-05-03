param(
    [string]$Version = "2.2.50"
)

$ErrorActionPreference = "Stop"

$rootPath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$toolRoot = Join-Path $rootPath (".tools\obfuscar\" + $Version)
$toolExe = Join-Path $toolRoot "tools\Obfuscar.Console.exe"
if (Test-Path $toolExe) {
    Write-Output $toolExe
    exit 0
}

$downloadRoot = Join-Path $rootPath ".tools\downloads"
$packagePath = Join-Path $downloadRoot ("obfuscar." + $Version + ".nupkg")
$zipPath = Join-Path $downloadRoot ("obfuscar." + $Version + ".zip")
$packageUrl = "https://www.nuget.org/api/v2/package/Obfuscar/" + $Version
$legacyPackagePath = Join-Path $rootPath ".codex-temp\obfuscar.$Version.nupkg"
$legacyExtractedTool = Join-Path $rootPath ".codex-temp\pkg\tools\Obfuscar.Console.exe"

if (Test-Path $legacyExtractedTool) {
    Write-Output $legacyExtractedTool
    exit 0
}

New-Item -ItemType Directory -Force -Path $downloadRoot | Out-Null
if (!(Test-Path $packagePath)) {
    if (Test-Path $legacyPackagePath) {
        Copy-Item -LiteralPath $legacyPackagePath -Destination $packagePath -Force
    }
    else {
        Invoke-WebRequest -Uri $packageUrl -OutFile $packagePath
    }
}

if (Test-Path $toolRoot) {
    Remove-Item -LiteralPath $toolRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $toolRoot | Out-Null
Copy-Item -LiteralPath $packagePath -Destination $zipPath -Force
Expand-Archive -LiteralPath $zipPath -DestinationPath $toolRoot -Force
Remove-Item -LiteralPath $zipPath -Force

if (!(Test-Path $toolExe)) {
    throw "Obfuscar.Console.exe not found after extraction."
}

Write-Output $toolExe
