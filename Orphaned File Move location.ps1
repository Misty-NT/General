# Search and list the Orphaned files in the installer folder and then move them to the NT folder. 
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script must be run as Administrator!"
    exit
}

# Set Source and Destination
$installerFolder = "$env:windir\Installer"
$destinationFolder = "C:\NetworkTitan\Installer Folder Cleanup"

# Create destination folder if it doesn't exist
if (-not (Test-Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
}

# Get list of all .msi and .msp files in the Installer folder
$installerFiles = Get-ChildItem -Path $installerFolder -Filter *.ms* -Recurse -File

# Get referenced products from registry
$referencedFiles = @()

# Check for product cache references (used by MSI)
$keyPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Patches"
)

foreach ($keyPath in $keyPaths) {
    Get-ChildItem -Path $keyPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $_.GetValueNames() | ForEach-Object {
            $val = $_.ToString()
            if ($val -match "\.ms(i|p)$") {
                $referencedFiles += $val.ToLower()
            }
        }
    }
}

# Build a list of orphaned files
$orphans = @()

foreach ($file in $installerFiles) {
    if ($referencedFiles -notcontains $file.Name.ToLower()) {
        $orphans += $file
    }
}

# Move orphaned files
foreach ($orphan in $orphans) {
    try {
        $targetPath = Join-Path -Path $destinationFolder -ChildPath $orphan.Name
        Move-Item -Path $orphan.FullName -Destination $targetPath -Force
        Write-Host "Moved: $($orphan.Name)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to move $($orphan.Name): $_"
    }
}

Write-Host "`n Completed. $($orphans.Count) orphaned file(s) moved to: $destinationFolder" -ForegroundColor Cyan

