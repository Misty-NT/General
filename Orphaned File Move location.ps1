# === Define Paths ===
$Folder = "C:\NetworkTitan\WICleanup"
$MoveDestination = "C:\NetworkTitan\Installer Folder Cleanup"
$LogPath = "$Folder\WICleanup-MoveLog.txt"

# === Ensure folders exist ===
foreach ($path in @($Folder, $MoveDestination)) {
    if (-Not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

# === Initialize Installer COM Object ===
$installer = New-Object -ComObject WindowsInstaller.Installer
$usedFiles = @{}

# === Collect in-use MSI/MSP files ===
foreach ($product in $installer.Products) {
    try {
        foreach ($patch in $installer.Patches($product)) {
            $patchPath = $installer.PatchInfo($patch, "LocalPackage")
            if ($patchPath) {
                $usedFiles[$patchPath.ToLower()] = $true
            }
        }

        $localPackage = $installer.ProductInfo($product, "LocalPackage")
        if ($localPackage) {
            $usedFiles[$localPackage.ToLower()] = $true
        }
    } catch {
        # Continue on error
    }
}

# === Scan C:\Windows\Installer for orphaned .msi/.msp files ===
$installerFolder = "C:\Windows\Installer"
$orphanedFiles = @()

Get-ChildItem -Path $installerFolder -File -Filter *.ms* | ForEach-Object {
    $lowerPath = $_.FullName.ToLower()
    if (-not $usedFiles.ContainsKey($lowerPath)) {
        $orphanedFiles += $_
    }
}

# === Move orphaned files and log ===
Add-Content -Path $LogPath -Value "=== Orphaned Installer Files Moved on $(Get-Date) ==="

foreach ($file in $orphanedFiles) {
    try {
        $destinationPath = Join-Path -Path $MoveDestination -ChildPath $file.Name
        Move-Item -Path $file.FullName -Destination $destinationPath -Force
        Add-Content -Path $LogPath -Value "Moved: $($file.FullName) → $destinationPath"
        Write-Host "Moved: $($file.Name)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to move $($file.Name): $_"
        Add-Content -Path $LogPath -Value "ERROR: Failed to move $($file.FullName): $_"
    }
}

Write-Host "`n📄 Log written to: $LogPath"
Write-Host "✔️ $($orphanedFiles.Count) orphaned file(s) were moved to: $MoveDestination"

