$sysmon = Get-Service -Name Sysmon -ErrorAction SilentlyContinue

if ($sysmon) {
    Stop-Service Sysmon -Force
    Start-Process -FilePath "sysmon.exe" -ArgumentList "-u" -Wait -NoNewWindow
}

# or 
# sysmon.exe -u force

