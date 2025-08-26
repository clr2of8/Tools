# Create Desktop Shortcut for Audio Service Restart
# This script creates a desktop shortcut that runs the audio service restart command with elevated privileges

# Get the desktop path
$DesktopPath = [Environment]::GetFolderPath("Desktop")

# Define the shortcut name
$ShortcutName = "Restart Audio Service.lnk"
$ShortcutPath = Join-Path $DesktopPath $ShortcutName

# Define the target command
$TargetCommand = "powershell.exe"
$Arguments = "-nop -exec bypass -File `"$env:USERPROFILE\Documents\code\Tools\Restart_Audio_Service\launch_restart_audio_service.ps1`""

# Create the shortcut object
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)

# Set shortcut properties
$Shortcut.TargetPath = $TargetCommand
$Shortcut.Arguments = $Arguments
$Shortcut.WorkingDirectory = "$env:USERPROFILE\Documents\code\Tools"
$Shortcut.Description = "Restart Audio Service with elevated privileges"
$Shortcut.IconLocation = "powershell.exe,0"

# Save the shortcut
$Shortcut.Save()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Desktop shortcut created successfully at: $ShortcutPath" -ForegroundColor Green
Write-Host "Shortcut name: $ShortcutName" -ForegroundColor Green
Write-Host "You can now double-click the shortcut on your desktop to restart the audio service." -ForegroundColor Yellow
