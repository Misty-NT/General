# Ensure the output folder exists
$folder = "C:\NetworkTitan"
if (-not (Test-Path $folder)) {
    New-Item -Path $folder -ItemType Directory | Out-Null
}

# Build timestamped file path
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$csvPath   = Join-Path $folder "Processes_$timestamp.csv"

# Collect processes and owners
$processList = Get-WmiObject Win32_Process | ForEach-Object {
    try {
        $owner = $_.GetOwner()
        if ($owner.User) {
            [PSCustomObject]@{
                Computer    = $env:COMPUTERNAME
                ProcessName = $_.Name
                PID         = $_.ProcessId
                User        = "$($owner.Domain)\$($owner.User)"
                CommandLine = $_.CommandLine
            }
        }
    } catch { }
}

# Export to CSV
$processList | Sort-Object User, ProcessName | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Process list saved to $csvPath"

# Retention cleanup (30 days)
Get-ChildItem $folder -Filter "Processes_*.csv" | Where-Object {
    $_.LastWriteTime -lt (Get-Date).AddDays(-30)
} | Remove-Item -Force -ErrorAction SilentlyContinue
