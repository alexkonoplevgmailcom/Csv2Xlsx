@echo off
setlocal

REM Check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This script requires administrator privileges.
    echo Right-click on the script and select "Run as administrator".
    pause
    exit /B 1
)

REM Set the service name
set SERVICE_NAME=Csv2XlsxService
set SERVICE_DISPLAY_NAME=CSV to XLSX Converter Service
set SERVICE_DESCRIPTION="Monitors a folder for CSV files and converts them to Excel XLSX files based on a template and mapping."
set BINARY_PATH=%~dp0Csv2Xlsx.exe --service

REM Get the command argument
set COMMAND=%1

if "%COMMAND%"=="" (
    echo Usage: ServiceManager.bat [install^|start^|stop^|uninstall]
    echo.
    echo Commands:
    echo   install   - Install the Windows Service
    echo   start     - Start the Windows Service
    echo   stop      - Stop the Windows Service
    echo   uninstall - Uninstall the Windows Service
    goto :EOF
)

if /i "%COMMAND%"=="install" (
    echo Installing %SERVICE_NAME% service...
    sc create %SERVICE_NAME% binPath= "%BINARY_PATH%" DisplayName= "%SERVICE_DISPLAY_NAME%" start= auto
    if %errorLevel% neq 0 (
        echo Failed to create service. Make sure you're running as administrator.
        exit /B 1
    )
    sc description %SERVICE_NAME% %SERVICE_DESCRIPTION%
    echo Service installed successfully.
    goto :EOF
)

if /i "%COMMAND%"=="start" (
    echo Starting %SERVICE_NAME% service...
    sc start %SERVICE_NAME%
    if %errorLevel% neq 0 (
        echo Failed to start service. Make sure the service is installed correctly.
        exit /B 1
    )
    echo Service start command issued.
    goto :EOF
)

if /i "%COMMAND%"=="stop" (
    echo Stopping %SERVICE_NAME% service...
    sc stop %SERVICE_NAME%
    if %errorLevel% neq 0 (
        echo Failed to stop service. Make sure the service is running.
        exit /B 1
    )
    echo Service stop command issued.
    goto :EOF
)

if /i "%COMMAND%"=="uninstall" (
    echo Uninstalling %SERVICE_NAME% service...
    sc stop %SERVICE_NAME%
    sc delete %SERVICE_NAME%
    echo Service uninstalled successfully.
    goto :EOF
)

echo Unknown command: %COMMAND%
echo Use 'ServiceManager.bat' without arguments to see usage.