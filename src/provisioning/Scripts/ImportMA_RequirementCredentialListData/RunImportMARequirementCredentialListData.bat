@echo off
REM Import MA Requirement Credential List Data Script
REM This script imports MA Requirement Credential data from CSV to SharePoint list

setlocal enabledelayedexpansion

REM Configuration
set "SITE_URL=https://eutelsatgroup.sharepoint.com/sites/MADB-Dev"
set "LIST_NAME=MA Requirement"
set "SCRIPT_DIR=%~dp0"
set "CSV_FILE=%SCRIPT_DIR%MA_RequirementCredentialData.csv"
set "LOG_DIR=%SCRIPT_DIR%Logs"
set "LOG_FILE=%LOG_DIR%\ImportMARequirementCredentialListData_%date:~-4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log"
set "BATCH_SIZE=2"
set "DRY_RUN=false"

REM Create log directory if it doesn't exist
if not exist "!LOG_DIR!" mkdir "!LOG_DIR!"

REM Execute PowerShell script
echo Importing MA Requirement Credential data to SharePoint...
echo Log file: !LOG_FILE!

if "!DRY_RUN!"=="true" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -Command ^
        "& '!SCRIPT_DIR!ImportMARequirementCredentialListData.ps1' -SiteUrl '!SITE_URL!' -ListName '!LIST_NAME!' -CsvFilePath '!CSV_FILE!' -LogFilePath '!LOG_FILE!' -BatchSize !BATCH_SIZE! -DryRun -Debug -Verbose"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -Command ^
        "& '!SCRIPT_DIR!ImportMARequirementCredentialListData.ps1' -SiteUrl '!SITE_URL!' -ListName '!LIST_NAME!' -CsvFilePath '!CSV_FILE!' -LogFilePath '!LOG_FILE!' -BatchSize !BATCH_SIZE! -Verbose"
)

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Import process completed successfully.
    echo Log file saved to: !LOG_FILE!
) else (
    echo.
    echo Import process failed with error code: %ERRORLEVEL%
    echo Check log file for details: !LOG_FILE!
    pause
)

endlocal
