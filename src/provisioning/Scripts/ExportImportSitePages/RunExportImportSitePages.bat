@echo off
REM Export/Import SharePoint Site Pages Script
REM Copies Site Pages from one site collection to another.

setlocal enabledelayedexpansion

REM Configuration
set "SOURCE_SITE_URL=https://eutelsatgroupnam.sharepoint.com/sites/BusinessOps"
set "TARGET_SITE_URL=https://eutelsatgroupnam.sharepoint.com/sites/SalesExcellenceScope"
set "SCRIPT_DIR=%~dp0"
set "LOG_DIR=%SCRIPT_DIR%Logs"
set "PAGE_NAMES="
set "BATCH_SIZE=10"
set "DRY_RUN=true"

REM Optional command-line overrides
REM Arg1 = SourceSiteUrl
REM Arg2 = TargetSiteUrl
REM Arg3 = PageNamesCsv (example: Home.aspx,Contact.aspx)
REM Arg4 = DryRun (true/false)
if not "%~1"=="" set "SOURCE_SITE_URL=%~1"
if not "%~2"=="" set "TARGET_SITE_URL=%~2"
if not "%~3"=="" set "PAGE_NAMES=%~3"
if not "%~4"=="" set "DRY_RUN=%~4"

if not exist "!LOG_DIR!" mkdir "!LOG_DIR!"

set "HH=%time:~0,2%"
if "!HH:~0,1!"==" " set "HH=0!HH:~1,1!"
set "LOG_FILE=!LOG_DIR!\ExportImportSitePages_%date:~-4%%date:~-10,2%%date:~-7,2%_!HH!!time:~3,2!!time:~6,2!.log"

echo Export/Import Site Pages

echo Source Site: !SOURCE_SITE_URL!
echo Target Site: !TARGET_SITE_URL!
echo Page Names: !PAGE_NAMES!
echo Dry Run: !DRY_RUN!
echo Log file: !LOG_FILE!

set "POWERSHELL_ARGS=-SourceSiteUrl '!SOURCE_SITE_URL!' -TargetSiteUrl '!TARGET_SITE_URL!' -LogFilePath '!LOG_FILE!' -BatchSize !BATCH_SIZE!"

if not "!PAGE_NAMES!"=="" (
    set "POWERSHELL_ARGS=!POWERSHELL_ARGS! -PageNamesCsv '!PAGE_NAMES!'"
)

if /I "!DRY_RUN!"=="true" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -Command ^
        "& '!SCRIPT_DIR!ExportImportSitePages.ps1' !POWERSHELL_ARGS! -DryRun -Debug -Verbose"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -Command ^
        "& '!SCRIPT_DIR!ExportImportSitePages.ps1' !POWERSHELL_ARGS! -Verbose"
)

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Export/Import process completed successfully.
    echo Log file saved to: !LOG_FILE!
) else (
    echo.
    echo Export/Import process failed with error code: %ERRORLEVEL%
    echo Check log file for details: !LOG_FILE!
    pause
)

endlocal
