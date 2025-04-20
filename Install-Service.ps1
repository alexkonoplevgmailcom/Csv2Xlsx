# Requires -RunAsAdministrator
# Script to install and manage the Csv2XlsxService

param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("install", "start", "stop", "uninstall")]
    [string]$Command
)

$ServiceName = "Csv2XlsxService"
$DisplayName = "CSV to XLSX Converter Service"
$Description = "Monitors a folder for CSV files and converts them to Excel XLSX files based on a template and mapping."
$BinaryPath = Join-Path $PSScriptRoot "Csv2Xlsx.exe --service"

function Install-Service {
    Write-Host "Installing $ServiceName service..."
    $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    
    if ($existingService -ne $null) {
        Write-Host "Service already exists. Removing it first..."
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        sc.exe delete $ServiceName | Out-Null
        Start-Sleep -Seconds 2
    }
    
    $result = sc.exe create $ServiceName binPath= "$BinaryPath" DisplayName= "$DisplayName" start= auto
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to create service: $result" -ForegroundColor Red
        exit 1
    }
    
    sc.exe description $ServiceName $Description
    Write-Host "Service installed successfully." -ForegroundColor Green
}

function Start-ServiceWithRetry {
    Write-Host "Starting $ServiceName service..."
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    
    if ($service -eq $null) {
        Write-Host "Service does not exist. Please install it first." -ForegroundColor Red
        exit 1
    }
    
    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $retryCount = 0
    
    while ((Get-Service -Name $ServiceName).Status -ne 'Running' -and $retryCount -lt 5) {
        Write-Host "Waiting for service to start..."
        Start-Sleep -Seconds 2
        $retryCount++
    }
    
    if ((Get-Service -Name $ServiceName).Status -eq 'Running') {
        Write-Host "Service started successfully." -ForegroundColor Green
    } else {
        Write-Host "Failed to start service. Check the Event Log for more details." -ForegroundColor Red
    }
}

function Stop-ServiceSafely {
    Write-Host "Stopping $ServiceName service..."
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    
    if ($service -eq $null) {
        Write-Host "Service does not exist." -ForegroundColor Yellow
        return
    }
    
    if ($service.Status -eq 'Running') {
        Stop-Service -Name $ServiceName -Force
        $retryCount = 0
        
        while ((Get-Service -Name $ServiceName).Status -ne 'Stopped' -and $retryCount -lt 5) {
            Write-Host "Waiting for service to stop..."
            Start-Sleep -Seconds 2
            $retryCount++
        }
        
        if ((Get-Service -Name $ServiceName).Status -eq 'Stopped') {
            Write-Host "Service stopped successfully." -ForegroundColor Green
        } else {
            Write-Host "Failed to stop service. It may still be processing requests." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Service is not running." -ForegroundColor Yellow
    }
}

function Uninstall-Service {
    Write-Host "Uninstalling $ServiceName service..."
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    
    if ($service -eq $null) {
        Write-Host "Service does not exist." -ForegroundColor Yellow
        return
    }
    
    Stop-ServiceSafely
    sc.exe delete $ServiceName | Out-Null
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Service uninstalled successfully." -ForegroundColor Green
    } else {
        Write-Host "Failed to uninstall service." -ForegroundColor Red
    }
}

# Execute the requested command
switch ($Command) {
    "install" { Install-Service }
    "start" { Start-ServiceWithRetry }
    "stop" { Stop-ServiceSafely }
    "uninstall" { Uninstall-Service }
}