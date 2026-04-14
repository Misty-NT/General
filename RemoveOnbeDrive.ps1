# =============================
# Remove OneDrive Completely
# =============================

Write-Host "Stopping OneDrive processes..."
taskkill /f /im OneDrive.exe 2>$null

# Detect system architecture
$SystemRoot = $env:SystemRoot

if (Test-Path "$SystemRoot\SysWOW64\OneDriveSetup.exe") {
    $OneDriveSetup = "$SystemRoot\SysWOW64\OneDriveSetup.exe"
} else {
    $OneDriveSetup = "$SystemRoot\System32\OneDriveSetup.exe"
}

Write-Host "Uninstalling OneDrive..."
Start-Process $OneDriveSetup "/uninstall" -NoNewWindow -Wait

# =============================
# Remove leftover folders
# =============================

Write-Host "Removing leftover files..."

$paths = @(
    "$env:LOCALAPPDATA\Microsoft\OneDrive",
    "$env:PROGRAMDATA\Microsoft OneDrive",
    "$env:SYSTEMDRIVE\OneDriveTemp",
    "$env:USERPROFILE\OneDrive"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# =============================
# Remove OneDrive from Explorer
# =============================

Write-Host "Removing OneDrive from Explorer sidebar..."

New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" `
    -Name "DisableFileSyncNGSC" -Value 1 -Type DWord

# Remove from Explorer Namespace
$regPaths = @(
    "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
    "HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
)

foreach ($regPath in $regPaths) {
    if (Test-Path $regPath) {
        Set-ItemProperty -Path $regPath -Name "System.IsPinnedToNameSpaceTree" -Value 0
    }
}

# =============================
# Disable OneDrive via GPO registry
# =============================

Write-Host "Blocking OneDrive reinstall..."

Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" `
    -Name "DisableFileSyncNGSC" -Value 1 -Type DWord

# =============================
# Remove OneDrive from Startup
# =============================

Write-Host "Cleaning startup entries..."

$startupReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path $startupReg -Name "OneDrive" -ErrorAction SilentlyContinue

# =============================
# Remove scheduled tasks (if present)
# =============================

Write-Host "Removing scheduled tasks..."

Get-ScheduledTask -TaskName "*OneDrive*" -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false

# =============================
# Done
# =============================

Write-Host "OneDrive has been removed and blocked from reinstall."
