#Requires -Version 5.1
#Requires -RunAsAdministrator

# Script to restart Windows Audio service with admin privileges

function Restart-AudioService {
    try {
        Write-Host "Attempting to restart Windows Audio service..." -ForegroundColor Yellow
        
        # Get the audio service
        $audioService = Get-Service -Name "Audiosrv" -ErrorAction Stop
        
        if ($audioService.Status -eq "Running") {
            Write-Host "Stopping Windows Audio service..." -ForegroundColor Yellow
            Stop-Service -Name "Audiosrv" -Force -ErrorAction Stop
            Start-Sleep -Seconds 2
            
            Write-Host "Starting Windows Audio service..." -ForegroundColor Yellow
            Start-Service -Name "Audiosrv" -ErrorAction Stop
            Start-Sleep -Seconds 2
            
            # Verify service is running
            $audioService = Get-Service -Name "Audiosrv"
            if ($audioService.Status -eq "Running") {
                Write-Host "Windows Audio service restarted successfully!" -ForegroundColor Green
                Write-Host "Service Status: $($audioService.Status)" -ForegroundColor Green
            } else {
                Write-Host "Warning: Audio service status is $($audioService.Status)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Windows Audio service is not currently running. Starting it..." -ForegroundColor Yellow
            Start-Service -Name "Audiosrv" -ErrorAction Stop
            Start-Sleep -Seconds 2
            
            $audioService = Get-Service -Name "Audiosrv"
            if ($audioService.Status -eq "Running") {
                Write-Host "Windows Audio service started successfully!" -ForegroundColor Green
            } else {
                Write-Host "Failed to start Windows Audio service. Status: $($audioService.Status)" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Make sure you have administrator privileges and the service exists." -ForegroundColor Red
        exit 1
    }
}

# Main execution
Write-Host "Windows Audio Service Restart Script" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

# Check if audio service exists
try {
    $audioService = Get-Service -Name "Audiosrv" -ErrorAction Stop
    Write-Host "Audio service found. Current status: $($audioService.Status)" -ForegroundColor Green
}
catch {
    Write-Host "Error: Windows Audio service (Audiosrv) not found." -ForegroundColor Red
    Write-Host "This script is designed for Windows systems with the standard audio service." -ForegroundColor Red
    exit 1
}

# Restart the audio service
Restart-AudioService

Write-Host "`nScript completed." -ForegroundColor Cyan
Sleep 2
