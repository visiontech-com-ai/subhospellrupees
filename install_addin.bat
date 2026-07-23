@echo off
setlocal enabledelayedexpansion

title SubhoSpellRupees - Excel Currency Add-In Installer
cls

echo ===================================================
echo   SubhoSpellRupees - Excel Add-In Installer
echo   Product  : Indian Currency Speller
echo   Developer: Subho
echo ===================================================
echo.

:: --- CONFIGURATION ---
set "ADDIN_NAME=SubhoSpellRupees.xlam"

set "GITHUB_RAW_URL=https://raw.githubusercontent.com/visiontech-com-ai/subhospellrupees/main/%ADDIN_NAME%"
set "ADDINS_DIR=%APPDATA%\Microsoft\AddIns"

:: Target path
set "DEST_FILE=%ADDINS_DIR%\%ADDIN_NAME%"

:: Create AddIns directory if it does not exist
if not exist "%ADDINS_DIR%" (
    echo [i] Creating Excel AddIns directory...
    mkdir "%ADDINS_DIR%"
)

echo [>] Fetching %ADDIN_NAME% from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " ^
    "try { " ^
    "    Invoke-WebRequest -Uri '%GITHUB_RAW_URL%' -OutFile '%DEST_FILE%' -UseBasicParsing; " ^
    "    Write-Host '[+] Download successful.' -ForegroundColor Green; " ^
    "} catch { " ^
    "    Write-Host '[-] Failed to download file. Please check repository URL or internet connection.' -ForegroundColor Red; " ^
    "    exit 1; " ^
    "}"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [-] Installation failed.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [>] Registering Add-In with Microsoft Excel...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { " ^
    "    $excel = New-Object -ComObject Excel.Application; " ^
    "    $excel.Visible = $false; " ^
    "    $addin = $excel.AddIns.Add('%DEST_FILE%', $false); " ^
    "    $addin.Installed = $true; " ^
    "    $excel.Quit(); " ^
    "    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null; " ^
    "    Write-Host '[+] Successfully enabled Add-In in Excel.' -ForegroundColor Green; " ^
    "} catch { " ^
    "    Write-Host '[!] Downloaded to AddIns folder, but COM registration requires manual activation via Excel -> Add-ins.' -ForegroundColor Yellow; " ^
    "}"

echo.
echo ===================================================
echo   Installation Complete!
echo   Function =SpellRupees() is now available system-wide.
echo   Powered by SubhoSpellRupees
echo ===================================================
echo.
pause