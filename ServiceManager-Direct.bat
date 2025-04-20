@echo off
setlocal enabledelayedexpansion

:menu
cls
echo CSV to XLSX Service Installer
echo ===========================
echo.
echo This script will install, start, stop, or uninstall the Csv2XlsxService.
echo.
echo Menu:
echo 1. Install Service
echo 2. Start Service
echo 3. Stop Service
echo 4. Uninstall Service
echo 5. Check Service Status
echo 6. View Log Files
echo 7. Exit
echo.

set /p option=Enter option (1-7): 

if "%option%"=="1" goto install
if "%option%"=="2" goto start
if "%option%"=="3" goto stop
if "%option%"=="4" goto uninstall
if "%option%"=="5" goto status
if "%option%"=="6" goto viewlogs
if "%option%"=="7" goto exit
goto menu

:install
echo.
echo Installing Csv2XlsxService...

rem Create logs directory
if not exist "%~dp0bin\Release\net8.0\logs" (
    mkdir "%~dp0bin\Release\net8.0\logs"
    echo Created logs directory.
)

sc create Csv2XlsxService binPath= "%~dp0bin\Release\net8.0\Csv2Xlsx.exe --service" DisplayName= "CSV to XLSX Converter Service" start= auto
sc description Csv2XlsxService "Monitors a folder for CSV files and converts them to Excel XLSX files based on a template and mapping."
echo.
echo Service installed. You can now start it with option 2.
call :ask_return_to_menu
goto :eof

:start
echo.
echo Starting Csv2XlsxService...
sc start Csv2XlsxService
call :ask_return_to_menu
goto :eof

:stop
echo.
echo Stopping Csv2XlsxService...
sc stop Csv2XlsxService
call :ask_return_to_menu
goto :eof

:uninstall
echo.
echo Uninstalling Csv2XlsxService...
sc stop Csv2XlsxService
sc delete Csv2XlsxService
call :ask_return_to_menu
goto :eof

:status
echo.
echo Checking Csv2XlsxService status...
sc query Csv2XlsxService
call :ask_return_to_menu
goto :eof

:viewlogs
echo.
echo Log files are located at: %~dp0bin\Release\net8.0\logs\
echo.
echo Opening logs folder...
start explorer "%~dp0bin\Release\net8.0\logs\"
call :ask_return_to_menu
goto :eof

:ask_return_to_menu
echo.
echo This window will close in 30 seconds if no choice is made...
echo Press 'Y' to return to menu or 'N' to exit.
choice /c YN /t 30 /d N /n > nul
if errorlevel 2 goto exit
goto menu

:exit
echo.
echo Exiting...
exit /b 0