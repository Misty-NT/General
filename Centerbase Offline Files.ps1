# Centerbase OfflineEmails Size Report
# Displays results directly in the PowerShell window

$Results = @()

Get-ChildItem "C:\Users" -Directory | ForEach-Object {

    $WindowsProfile = $_.Name
    $CenterbaseRoot = Join-Path $_.FullName "AppData\Roaming\Centerbase"

    if (Test-Path $CenterbaseRoot) {

        Get-ChildItem $CenterbaseRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {

            $CenterbaseUser = $_.Name
            $OfflineEmailPath = Join-Path $_.FullName "OfflineEmails"

            if (Test-Path $OfflineEmailPath) {

                $Files = Get-ChildItem $OfflineEmailPath -File -Recurse -Force -ErrorAction SilentlyContinue
                $SizeBytes = ($Files | Measure-Object Length -Sum).Sum

                $Results += [PSCustomObject]@{
                    Profile         = $WindowsProfile
                    CenterbaseUser  = $CenterbaseUser
                    Files           = $Files.Count
                    SizeGB          = [math]::Round($SizeBytes / 1GB, 2)
                    Folder          = $OfflineEmailPath
                }
            }
        }
    }
}

if ($Results.Count -gt 0) {
    $Results |
        Sort-Object SizeGB -Descending |
        Format-Table Profile, CenterbaseUser, Files, SizeGB, Folder -AutoSize
}
else {
    Write-Host "No Centerbase OfflineEmails folders were found."
}
