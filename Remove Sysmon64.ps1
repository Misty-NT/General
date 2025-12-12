# Remove Sysmon (64-bit)
$sysmon = Get-Service -Name Sysmon64 -ErrorAction SilentlyContinue

if ($sysmon) {
    Write-Output "Sysmon service found. Stopping and uninstalling..."
    Stop-Service Sysmon64 -Force
    Start-Process -FilePath "sysmon64.exe" -ArgumentList "-u" -Wait -NoNewWindow
    Write-Output "Sysmon removed successfully."
} else {
    Write-Output "Sysmon is not installed."
}
