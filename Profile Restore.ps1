# Restore All  

$printerCsv = "C:\NetworkTitan\Profile\EndpointBackup_20250827_185940\Printers\Printers.csv"
$printers = Import-Csv $printerCsv

foreach ($p in $printers) {
    try {
        # Ensure the driver is available before adding printer
        if (-not (Get-PrinterDriver -Name $p.DriverName -ErrorAction SilentlyContinue)) {
            Write-Warning "Driver $($p.DriverName) missing, install driver before restoring $($p.Name)."
            continue
        }

        # Add printer back (PortName must exist)
        if (-not (Get-Printer -Name $p.Name -ErrorAction SilentlyContinue)) {
            Add-Printer -Name $p.Name -DriverName $p.DriverName -PortName $p.PortName
            Write-Host "Restored printer: $($p.Name)"
        } else {
            Write-Host "Printer $($p.Name) already exists."
        }
    }
    catch {
        Write-Warning "Failed to restore printer $($p.Name): $_"
    }
}



# ---------------------------------------------------
# Option 2
# Driver & Port 

 $ports = Import-Csv "...\Printers\PrinterPorts.csv"
foreach ($port in $ports) {
    if (-not (Get-PrinterPort -Name $port.Name -ErrorAction SilentlyContinue)) {
        Add-PrinterPort -Name $port.Name -PrinterHostAddress $port.PrinterHostAddress
    }
}

-----------------------------------------------------
# Option 3 
#default Printer 
$defaultJson = Get-Content "...\Printers\DefaultPrinter_System.json" | ConvertFrom-Json
$defaultName = $defaultJson.DefaultPrinterName_GetPrinter
if ($defaultName) {
    Set-Printer -Name $defaultName -IsDefault $true
    Write-Host "Set default printer to $defaultName"
}
