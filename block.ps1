# Disable USB Storage Access
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR"

# Ensure the key exists
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

# Set Start value to 4 (Disabled)
Set-ItemProperty -Path $regPath -Name "Start" -Value 4

Write-Output "USB storage access has been disabled."
