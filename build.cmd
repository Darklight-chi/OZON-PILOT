@echo off
setlocal

set ROOT=%~dp0
set MSBUILD=C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe
set CONFIGURATION=Release
set BUILDROOT=%ROOT%.build
set OBFUSCAR_VERSION=2.2.50
if not exist "%MSBUILD%" set MSBUILD=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe

taskkill /im OZON-PILOT.exe /f >nul 2>nul
taskkill /im OZON-PILOT-Updater.exe /f >nul 2>nul
taskkill /im LitchiOzonRecovery.exe /f >nul 2>nul
taskkill /im LitchiAutoUpdate.exe /f >nul 2>nul

if not exist "%MSBUILD%" (
  echo MSBuild not found.
  exit /b 1
)

if not exist "%ROOT%build" mkdir "%ROOT%build"

echo Preparing obfuscator...
for /f "usebackq delims=" %%I in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%build\Ensure-Obfuscar.ps1" -Version "%OBFUSCAR_VERSION%"`) do set OBFUSCAR=%%I
if not exist "%OBFUSCAR%" (
  echo Obfuscar not found.
  exit /b 1
)

echo Building updater...
"%MSBUILD%" "%ROOT%src\LitchiAutoUpdate\LitchiAutoUpdate.csproj" /t:Build /p:Configuration=%CONFIGURATION% /verbosity:minimal
if errorlevel 1 exit /b 1

echo Building main app...
"%MSBUILD%" "%ROOT%src\LitchiOzonRecovery\LitchiOzonRecovery.csproj" /t:Build /p:Configuration=%CONFIGURATION% /verbosity:minimal
if errorlevel 1 exit /b 1

set MAINBIN=%ROOT%src\LitchiOzonRecovery\bin\%CONFIGURATION%
set UPDATERBIN=%ROOT%src\LitchiAutoUpdate\bin\%CONFIGURATION%
set OBFMAIN=%BUILDROOT%\obfuscated\main
set OBFUPDATER=%BUILDROOT%\obfuscated\updater

echo Obfuscating main app...
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%build\Invoke-Obfuscation.ps1" -ObfuscarExe "%OBFUSCAR%" -InputDirectory "%MAINBIN%" -OutputDirectory "%OBFMAIN%" -Modules "OZON-PILOT.exe"
if errorlevel 1 exit /b 1

echo Obfuscating updater...
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%build\Invoke-Obfuscation.ps1" -ObfuscarExe "%OBFUSCAR%" -InputDirectory "%UPDATERBIN%" -OutputDirectory "%OBFUPDATER%" -Modules "OZON-PILOT-Updater.exe"
if errorlevel 1 exit /b 1

set DIST=%ROOT%dist\OZON-PILOT
if exist "%DIST%" (
  rmdir /s /q "%DIST%" >nul 2>nul
  if exist "%DIST%" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$path = [System.IO.Path]::GetFullPath('%DIST%'); if (Test-Path -LiteralPath $path) { Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue; if ((Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0) { exit 1 } }"
    if errorlevel 1 (
      echo Failed to clean output folder: %DIST%
      exit /b 1
    )
  )
)
if not exist "%DIST%" mkdir "%DIST%"

echo Copying build output...
xcopy "%OBFMAIN%\OZON-PILOT.exe" "%DIST%\" /y /i /q >nul
xcopy "%MAINBIN%\Microsoft.Web.WebView2.Core.dll" "%DIST%\" /y /i /q >nul
xcopy "%MAINBIN%\Microsoft.Web.WebView2.WinForms.dll" "%DIST%\" /y /i /q >nul
xcopy "%ROOT%baseline\runtimes\win-x64\native\WebView2Loader.dll" "%MAINBIN%\runtimes\win-x64\native\" /y /i /q >nul
xcopy "%ROOT%baseline\runtimes\win-x86\native\WebView2Loader.dll" "%MAINBIN%\runtimes\win-x86\native\" /y /i /q >nul
xcopy "%ROOT%baseline\runtimes\win-x64\native\WebView2Loader.dll" "%MAINBIN%\" /y /i /q >nul
xcopy "%MAINBIN%\Newtonsoft.Json.dll" "%DIST%\" /y /i /q >nul
xcopy "%MAINBIN%\NPOI*.dll" "%DIST%\" /y /i /q >nul
 xcopy "%OBFUPDATER%\OZON-PILOT-Updater.exe" "%DIST%\" /y /i /q >nul
xcopy "%UPDATERBIN%\ICSharpCode.SharpZipLib.dll" "%DIST%\" /y /i /q >nul
xcopy "%MAINBIN%\zh-Hans\*.*" "%DIST%\zh-Hans\" /y /i /q >nul
xcopy "%ROOT%baseline\runtimes\win-x64\native\WebView2Loader.dll" "%DIST%\runtimes\win-x64\native\" /y /i /q >nul
xcopy "%ROOT%baseline\runtimes\win-x86\native\WebView2Loader.dll" "%DIST%\runtimes\win-x86\native\" /y /i /q >nul
xcopy "%ROOT%baseline\runtimes\win-x64\native\WebView2Loader.dll" "%DIST%\" /y /i /q >nul
 xcopy "%ROOT%baseline\*.*" "%DIST%\" /e /i /y /q >nul

 if exist "%DIST%\mscorlib.dll" del /q "%DIST%\mscorlib.dll"
 if exist "%DIST%\normidna.nlp" del /q "%DIST%\normidna.nlp"
 if exist "%DIST%\normnfc.nlp" del /q "%DIST%\normnfc.nlp"
 if exist "%DIST%\normnfd.nlp" del /q "%DIST%\normnfd.nlp"
 if exist "%DIST%\normnfkc.nlp" del /q "%DIST%\normnfkc.nlp"
 if exist "%DIST%\normnfkd.nlp" del /q "%DIST%\normnfkd.nlp"
 if exist "%DIST%\System.Data.SQLite.dll" del /q "%DIST%\System.Data.SQLite.dll"
 if exist "%DIST%\x64\SQLite.Interop.dll" del /q "%DIST%\x64\SQLite.Interop.dll"
 if exist "%DIST%\x86\SQLite.Interop.dll" del /q "%DIST%\x86\SQLite.Interop.dll"
 if exist "%DIST%\Data\SkuDb.db" del /q "%DIST%\Data\SkuDb.db"
 for /r "%DIST%" %%F in (*.pdb) do del /q "%%F" >nul 2>nul

echo Build complete.
echo Output: %DIST%
endlocal
