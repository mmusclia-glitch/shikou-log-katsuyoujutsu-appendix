@echo off
setlocal EnableExtensions DisableDelayedExpansion

set "SCRIPT_PATH="

if exist "%~dp0IndexPickup_v5_0_0.ps1" set "SCRIPT_PATH=%~dp0IndexPickup_v5_0_0.ps1"

if not defined SCRIPT_PATH (
    for /f "delims=" %%F in ('dir /b /a-d "%~dp0IndexPickup_v5_0_0*.ps1" 2^>nul') do (
        if not defined SCRIPT_PATH set "SCRIPT_PATH=%~dp0%%F"
    )
)

if not defined SCRIPT_PATH (
    for /f "delims=" %%F in ('dir /b /a-d "%~dp0IndexPickup_v5_*.ps1" 2^>nul') do (
        if not defined SCRIPT_PATH set "SCRIPT_PATH=%~dp0%%F"
    )
)

if not defined SCRIPT_PATH (
    echo ERROR: PowerShell script was not found in the same folder.
    echo Expected: IndexPickup_v5_0_0.ps1
    echo Folder  : %~dp0
    set "EXIT_CODE=9009"
    goto END
)

echo =============================================
echo Index Pickup Runner v5.0.0
echo =============================================
echo.

:ASK_SOURCE
set "SOURCE="
set /p "SOURCE=SOURCE="
if not defined SOURCE (
    echo SOURCE is empty.
    goto ASK_SOURCE
)

:ASK_SINCE
set "SINCE="
set /p "SINCE=SINCE (yyyy-MM-dd HH:mm)="
if not defined SINCE (
    echo SINCE is empty.
    goto ASK_SINCE
)

echo.
echo SOURCE  = %SOURCE%
echo SINCE   = %SINCE%
echo.
pause

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Source "%SOURCE%" -Since "%SINCE%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo ERROR: processing failed. ExitCode=%EXIT_CODE%
) else (
    echo DONE
    echo Stage 1 for ChatGPT:
    echo   1. Upload pickup_manifest.csv.
    echo   2. Upload RUN card.
    echo Then wait for the prompt instruction.
    echo After that, upload all files listed in RUN card [FILES].
)

:END
pause
endlocal & exit /b %EXIT_CODE%
