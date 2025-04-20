# Csv2XlsxService Manager
# This file provides easy-to-use options for managing the Windows Service

$host.UI.RawUI.WindowTitle = "Csv2XlsxService Manager"
Clear-Host
Write-Host "CSV to XLSX Converter Service Manager" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script requires administrator privileges to manage the service."
Write-Host ""
Write-Host "Select an option:" -ForegroundColor Yellow
Write-Host "1. Install service"
Write-Host "2. Start service"
Write-Host "3. Stop service"
Write-Host "4. Uninstall service"
Write-Host "5. Exit"
Write-Host ""

$option = Read-Host "Enter option (1-5)"

switch ($option) {
    "1" {
        Write-Host "Installing service..."
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\Install-Service.ps1`" -Command install" -Verb RunAs
    }
    "2" {
        Write-Host "Starting service..."
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\Install-Service.ps1`" -Command start" -Verb RunAs
    }
    "3" {
        Write-Host "Stopping service..."
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\Install-Service.ps1`" -Command stop" -Verb RunAs
    }
    "4" {
        Write-Host "Uninstalling service..."
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\Install-Service.ps1`" -Command uninstall" -Verb RunAs
    }
    "5" {
        Write-Host "Exiting..."
        exit
    }
    default {
        Write-Host "Invalid option. Exiting..." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

Write-Host ""
Write-Host "Operation initiated. Check the elevated PowerShell window for results."
Write-Host "Press any key to exit..."
$null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")