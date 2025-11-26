$Log = "C:\ProgramData\Ninja\dotnet-runtime-cleanup.log"
New-Item -ItemType File -Path $Log -Force | Out-Null

function Log {
    param ($Msg)
    Write-Output $Msg
    Add-Content $Log "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Msg"
}

Log "=== .NET Runtime Cleanup Started ==="

$AllowedToRemove = @("6.0.36")
$RuntimeName = "Microsoft Windows Desktop Runtime"

# Pull uninstall entries
$Installed = Get-ItemProperty `
  HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
  HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* `
  -ErrorAction SilentlyContinue |
Where-Object {
    $_.DisplayName -like "*Windows Desktop Runtime*"
} | Select DisplayName, UninstallString

if (-not $Installed) {
    Log "No Windows Desktop Runtimes found."
    exit 0
}

# Parse REAL version from DisplayName
$Parsed = foreach ($App in $Installed) {
    if ($App.DisplayName -match "Runtime\s-\s(\d+\.\d+\.\d+)") {
        [PSCustomObject]@{
            DisplayName = $App.DisplayName
            Version     = [version]$Matches[1]
            VersionText = $Matches[1]
            UninstallString = $App.UninstallString
        }
    }
}

if (-not $Parsed) {
    Log "No valid runtime versions parsed from DisplayName."
    exit 0
}

# Determine highest REAL runtime
$Highest = ($Parsed | Sort-Object Version -Descending | Select-Object -First 1)
Log "Highest valid runtime detected: $($Highest.VersionText)"

foreach ($App in $Parsed) {

    if ($AllowedToRemove -contains $App.VersionText) {

        if ($App.Version -ge $Highest.Version) {
            Log "SKIP (newest runtime): $($App.DisplayName)"
            continue
        }

        Log "Removing approved older runtime: $($App.DisplayName)"

        if ($App.UninstallString -match "\{[A-F0-9\-]+\}") {
            $Guid = $Matches[0]
            Log "Uninstall GUID: $Guid"

            Start-Process "msiexec.exe" `
                -ArgumentList "/x $Guid /qn /norestart" `
                -Wait `
                -NoNewWindow
        }
        else {
            Log "ERROR: No GUID found — skipping"
        }
    }
    else {
        Log "KEEP: $($App.DisplayName)"
    }
}

Log "=== Runtime cleanup completed ==="
exit 0
