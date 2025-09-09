# Set up logging
$logPath = "C:\NetworkTitan\Logs"
$logFile = "$logPath\ShortcutLog.txt"
if (!(Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath -Force }

Function Log {
    param ($msg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $msg"
}

# Start logging
Log "---- Starting shortcut creation ----"

# Use the correct OneDrive Desktop path
$targetUsername = "user"
$desktopPath = "C:\Users\$targetUsername\OneDrive - domain\Desktop"

if (!(Test-Path $desktopPath)) {
    Log "ERROR: OneDrive Desktop path '$desktopPath' does not exist."
} else {
    try {
        $shortcutName = "name.lnk"
        $targetPath = "\\path\to\shortcut\folder"
        $iconPath = "C:\Windows\System32\shell32.dll"
        $iconIndex = 4
        $shortcutFullPath = Join-Path -Path $desktopPath -ChildPath $shortcutName

        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($shortcutFullPath)
        $Shortcut.TargetPath = $targetPath
        $Shortcut.WorkingDirectory = $targetPath
        $Shortcut.IconLocation = "$iconPath,$iconIndex"
        $Shortcut.Save()

        Log "SUCCESS: Shortcut created at $shortcutFullPath pointing to $targetPath"
    }
    catch {
        Log "ERROR: Exception occurred - $_"
    }
}

Log "---- Script finished on $env:COMPUTERNAME ----"



-----------------------------------STOP--------------------------------------------------

# Confirm the shortcut 
$targetUsername = "user"
$shortcutPath = "C:\Users\$targetUsername\Desktop\name.lnk"

if (Test-Path $shortcutPath) {
    Write-Host "Shortcut EXISTS at: $shortcutPath"
} else {
    Write-Host "Shortcut NOT FOUND at: $shortcutPath"
}

