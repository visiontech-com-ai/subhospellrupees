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
    "$dest = '%DEST_FILE%'; " ^
    "$versions = @('16.0','15.0','14.0'); " ^
    "$optKey = $null; " ^
    "foreach ($v in $versions) { $k = 'HKCU:\SOFTWARE\Microsoft\Office\' + $v + '\Excel'; if (Test-Path $k) { $optKey = $k + '\Options'; if (-not (Test-Path $optKey)) { New-Item -Path $optKey -Force | Out-Null }; break } }; " ^
    "if (-not $optKey) { Write-Host '[!] Excel not found in registry. Activate manually: Excel > File > Options > Add-ins.' -ForegroundColor Yellow; exit 0 }; " ^
    "$props = Get-ItemProperty -Path $optKey -EA SilentlyContinue; " ^
    "$already = $false; " ^
    "if ($props) { $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN' } | ForEach-Object { if ($_.Value -like '*SubhoSpellRupees*') { $already = $true } } }; " ^
    "if ($already) { Write-Host '[+] Add-In is already registered.' -ForegroundColor Green; exit 0 }; " ^
    "$idx = ''; " ^
    "while (Get-ItemProperty -Path $optKey -Name ('OPEN' + $idx) -EA SilentlyContinue) { if ($idx -eq '') { $idx = '1' } else { $idx = [string]([int]$idx + 1) } }; " ^
    "$val = '/R ' + [char]34 + $dest + [char]34; " ^
    "Set-ItemProperty -Path $optKey -Name ('OPEN' + $idx) -Value $val -Type String; " ^
    "Write-Host '[+] Add-In registered. It will load automatically next time Excel opens.' -ForegroundColor Green"

echo.
echo ===================================================
echo   Installation Complete!
echo   Function =SpellRupees() is now available system-wide.
echo   If Excel is open, close and reopen it to activate.
echo   Powered by SubhoSpellRupees
echo ===================================================
echo.
pause