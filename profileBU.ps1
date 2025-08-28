<# 
.SYNOPSIS
  Exports printers, printer prefs, Bluetooth, Wi-Fi profiles, USB storage history, and mapped drives
  so you can restore/reference them if a local→domain profile switch goes sideways.

.PARAMETER TargetProfilePath
  Full path to the LOCAL user profile you’re converting (e.g., C:\Users\jdoe).
  Used to load that user’s HKCU hive to capture per-user printer connections, default printer, mapped drives, etc.

.PARAMETER OutputRoot
  Folder to write output. Default: C:\Temp\EndpointBackup_<timestamp>

.PARAMETER ExportWifiXml
  Also export Wi-Fi profiles as XML using netsh (optionally with clear keys via -WifiIncludeKeys).

.PARAMETER WifiIncludeKeys
  Include plaintext Wi-Fi passphrases in the exported XML (sensitive!). Requires -ExportWifiXml.

.EXAMPLE
  .\Export-EndpointUserConfig.ps1 -TargetProfilePath "C:\Users\scheduling" -ExportWifiXml

.EXAMPLE
  .\Export-EndpointUserConfig.ps1 -TargetProfilePath "C:\Users\scheduling" -ExportWifiXml -WifiIncludeKeys
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidateScript({Test-Path $_ -PathType Container})]
  [string]$TargetProfilePath,

  [string]$OutputRoot,

  [switch]$ExportWifiXml,
  [switch]$WifiIncludeKeys
)

function New-OutDir {
  param([string]$Path)
  if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Safe-ExportJson {
  param([Parameter(ValueFromPipeline=$true)]$Object, [string]$Path)
  process {
    try { $Object | ConvertTo-Json -Depth 6 | Out-File -FilePath $Path -Encoding UTF8 }
    catch {
      ($_ | Out-String) | Out-File -FilePath $Path -Encoding UTF8
    }
  }
}

# --- Prep output folders
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not $OutputRoot) { $OutputRoot = "C:\NetworkTitan\Profile\EndpointBackup_$stamp" }
New-OutDir $OutputRoot

$dirs = @{
  Printers   = Join-Path $OutputRoot "Printers"
  Bluetooth  = Join-Path $OutputRoot "Bluetooth"
  WiFi       = Join-Path $OutputRoot "WiFi"
  USB        = Join-Path $OutputRoot "USB"
  Drives     = Join-Path $OutputRoot "Drives"
  RegExports = Join-Path $OutputRoot "RegistryExports"
  Logs       = Join-Path $OutputRoot "Logs"
}
$dirs.GetEnumerator() | ForEach-Object { New-OutDir $_.Value }

$logFile = Join-Path $dirs.Logs "run.log"
Start-Transcript -Path $logFile -Force | Out-Null

Write-Host "`n=== Export Endpoint User Config (pre-domain-migration backup) ===`n"
Write-Host "Target profile : $TargetProfilePath"
Write-Host "Output folder  : $OutputRoot`n"

# -------------------------
# SECTION: PRINTERS (system-wide)
# -------------------------
Write-Host "Collecting printer objects..."
$printerList = @()
try { $printerList = Get-Printer -ErrorAction Stop } catch { Write-Warning "Get-Printer failed: $($_.Exception.Message)" }
$printerDrivers = @()
try { $printerDrivers = Get-PrinterDriver -ErrorAction Stop } catch { Write-Warning "Get-PrinterDriver failed: $($_.Exception.Message)" }
$printerPorts = @()
try { $printerPorts = Get-PrinterPort -ErrorAction Stop } catch { Write-Warning "Get-PrinterPort failed: $($_.Exception.Message)" }

# Per-printer configuration (best-effort; some drivers don’t expose)
$printConfigs = foreach ($p in $printerList) {
  try {
    $cfg = Get-PrintConfiguration -PrinterName $p.Name -ErrorAction Stop
    [pscustomobject]@{ PrinterName = $p.Name; Config = $cfg }
  } catch {
    [pscustomobject]@{ PrinterName = $p.Name; Config = $null; Note = "Get-PrintConfiguration failed: $($_.Exception.Message)"}
  }
}

$printerList  | Export-Csv (Join-Path $dirs.Printers "Printers.csv") -NoTypeInformation -Encoding UTF8
$printerDrivers | Export-Csv (Join-Path $dirs.Printers "PrinterDrivers.csv") -NoTypeInformation -Encoding UTF8
$printerPorts | Export-Csv (Join-Path $dirs.Printers "PrinterPorts.csv") -NoTypeInformation -Encoding UTF8
$printConfigs | Safe-ExportJson -Path (Join-Path $dirs.Printers "PrinterConfigs.json")

# Default printer (system view via WMI and Get-Printer)
$defaultPrinterObj = $null
try {
  $defaultPrinterObj = Get-Printer | Where-Object { $_.Default -eq $true }
} catch {}
$defaultPrinterWmi = $null
try {
  $defaultPrinterWmi = Get-CimInstance -Class Win32_Printer | Where-Object { $_.Default -eq $true }
} catch {}
[pscustomobject]@{
  DefaultPrinterName_GetPrinter = $defaultPrinterObj.Name
  DefaultPrinterName_WMI        = $defaultPrinterWmi.Name
} | Safe-ExportJson -Path (Join-Path $dirs.Printers "DefaultPrinter_System.json")

# -------------------------
# SECTION: PER-USER PRINTER CONNECTIONS & PREFS (load HKCU of target user)
# -------------------------
Write-Host "Loading target user's HKCU hive to capture per-user printer settings..."
$ntUserDat = Join-Path $TargetProfilePath "NTUSER.DAT"
if (-not (Test-Path $ntUserDat)) { Write-Warning "NTUSER.DAT not found at $ntUserDat. Skipping per-user registry exports." }
else {
  $tempHive = "HKU\TempUserHive_$stamp"
  & reg.exe load $tempHive $ntUserDat | Out-Null
  try {
    $userRegRoot = "Registry::" + $tempHive
    $userPrintersKey = Join-Path $userRegRoot "Software\Microsoft\Windows NT\CurrentVersion\Printers"
    $userWindowsKey  = Join-Path $userRegRoot "Software\Microsoft\Windows NT\CurrentVersion\Windows"

    # Export raw .reg copies for completeness (easy restore reference)
    & reg.exe export "$tempHive\Software\Microsoft\Windows NT\CurrentVersion\Printers" (Join-Path $dirs.RegExports "User_PrinterKeys.reg") /y | Out-Null
    & reg.exe export "$tempHive\Software\Microsoft\Windows NT\CurrentVersion\Windows"  (Join-Path $dirs.RegExports "User_WindowsKey.reg") /y | Out-Null

    # Read useful subkeys: Connections (network printers), Devices & PrinterPorts (legacy mappings), Settings (per-printer prefs)
    $connections = $null; $devices = $null; $printerPorts = $null; $settings = $null
    try { $connections = Get-Item -Path (Join-Path $userPrintersKey "Connections") -ErrorAction Stop | Get-ChildItem -Recurse -ErrorAction SilentlyContinue } catch {}
    try { $devices     = Get-ItemProperty -Path (Join-Path $userWindowsKey  "Devices") -ErrorAction SilentlyContinue } catch {}
    try { $printerPorts= Get-ItemProperty -Path (Join-Path $userWindowsKey  "PrinterPorts") -ErrorAction SilentlyContinue } catch {}
    try { $settings    = Get-ChildItem -Path (Join-Path $userPrintersKey "Settings") -ErrorAction SilentlyContinue | ForEach-Object {
                           [pscustomobject]@{ Printer = $_.PSChildName; Values = (Get-ItemProperty -Path $_.PSPath) }
                         }
         } catch {}

    # Per-user default printer (Windows key "Device")
    $defaultPerUser = $null
    try { $defaultPerUser = (Get-ItemProperty -Path $userWindowsKey -Name "Device" -ErrorAction Stop).Device } catch {}

    [pscustomobject]@{
      DefaultPrinter_PerUser = $defaultPerUser
      Connections_Subkeys    = $connections | Select-Object PSChildName, PSPath
      Devices_ValueBag       = $devices
      PrinterPorts_ValueBag  = $printerPorts
      Settings_PerPrinter    = $settings
    } | Safe-ExportJson -Path (Join-Path $dirs.Printers "PerUser_PrinterProfile.json")
  }
  finally {
    & reg.exe unload $tempHive | Out-Null
  }
}

# -------------------------
# SECTION: BLUETOOTH (paired device inventory)
# -------------------------
Write-Host "Collecting Bluetooth paired devices..."
$btOut = @()

# Device inventory via PnP
try {
  $btPnp = Get-PnpDevice -Class Bluetooth -Status OK -ErrorAction Stop
  $btOut += $btPnp | Select-Object FriendlyName, InstanceId, Present, Status, Problem
} catch {
  Write-Warning "Get-PnpDevice (Bluetooth) failed: $($_.Exception.Message)"
}

# Registry: paired device MACs (BTHPORT)
try {
  $bthKey = "HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Devices"
  if (Test-Path $bthKey) {
    $btReg = Get-ChildItem $bthKey -ErrorAction SilentlyContinue | ForEach-Object {
      $macKey = $_.PSChildName
      $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
      # Name may exist in subvalues; capture raw bag to be safe
      [pscustomobject]@{
        Source      = "Registry"
        MacKey      = $macKey
        Properties  = $props
      }
    }
    $btOut += $btReg
  }
} catch {
  Write-Warning "Bluetooth registry scrape failed: $($_.Exception.Message)"
}

$btOut | Safe-ExportJson -Path (Join-Path $dirs.Bluetooth "Bluetooth_Inventory.json")

# -------------------------
# SECTION: WI-FI PROFILES (optional)
# -------------------------
if ($ExportWifiXml) {
  Write-Host "Exporting Wi-Fi profiles via netsh (XMLs)..."
  $opt = "key=clear"
  if (-not $WifiIncludeKeys) { $opt = "" }
  # List profiles first for a summary
  $profilesTxt = Join-Path $dirs.WiFi "Profiles.txt"
  cmd.exe /c "netsh wlan show profiles > `"$profilesTxt`""

  # Export each profile XML (netsh handles filenames)
  $optPart = $opt
  if ($optPart) { $optPart = " $optPart" }
  cmd.exe /c "netsh wlan export profile folder=`"$($dirs.WiFi)`"$optPart" | more" | Out-Null

  if ($WifiIncludeKeys) {
    Out-File -FilePath (Join-Path $dirs.WiFi "WARNING.txt") -Encoding UTF8 -InputObject @"
These Wi-Fi profile XMLs include CLEAR KEYS (passwords).
Treat as sensitive data and delete after use.
"@
  }
}

# -------------------------
# SECTION: USB STORAGE HISTORY
# -------------------------
Write-Host "Collecting USB storage history..."
$usbList = @()
try {
  $usbEnumKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR"
  if (Test-Path $usbEnumKey) {
    $usbList = Get-ChildItem $usbEnumKey -ErrorAction SilentlyContinue | ForEach-Object {
      $devName = $_.PSChildName
      Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue | ForEach-Object {
        $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{
          Device      = $devName
          Instance    = $_.PSChildName
          Friendly    = $props.FriendlyName
          Mfg         = $props.Mfg
          ParentId    = $props.ParentIdPrefix
          Class       = $props.Class
        }
      }
    }
  }
} catch {
  Write-Warning "USB history read failed: $($_.Exception.Message)"
}
$usbList | Export-Csv (Join-Path $dirs.USB "USBSTOR.csv") -NoTypeInformation -Encoding UTF8

# -------------------------
# SECTION: MAPPED DRIVES (current + per-user via hive)
# -------------------------
Write-Host "Collecting mapped drive info..."
$currentMapped = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -like '\\*' } | 
                 Select-Object Name, Root, Used, Free, Description
$currentMapped | Export-Csv (Join-Path $dirs.Drives "MappedDrives_CurrentSession.csv") -NoTypeInformation -Encoding UTF8

# Per-user (from loaded hive earlier) – if we still have the path, reload briefly to read Network & MountPoints2
if (Test-Path $ntUserDat) {
  $tempHive2 = "HKU\TempUserHive2_$stamp"
  & reg.exe load $tempHive2 $ntUserDat | Out-Null
  try {
    $root = "Registry::" + $tempHive2
    $netKey = Join-Path $root "Network"
    $mpKey  = Join-Path $root "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
    $mappedPerUser = @()

    if (Test-Path $netKey) {
      $mappedPerUser += Get-ChildItem $netKey -ErrorAction SilentlyContinue | ForEach-Object {
        $p = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{
          DriveLetter = $_.PSChildName
          RemotePath  = $p.RemotePath
          UserName    = $p.UserName
          Provider    = $p.ProviderName
          DeferFlags  = $p.DeferFlags
        }
      }
    }

    $mountPoints = $null
    if (Test-Path $mpKey) {
      $mountPoints = Get-ChildItem $mpKey -ErrorAction SilentlyContinue | Select-Object PSChildName, PSPath
    }

    $mappedPerUser | Export-Csv (Join-Path $dirs.Drives "MappedDrives_PerUser.csv") -NoTypeInformation -Encoding UTF8
    $mountPoints  | Safe-ExportJson -Path (Join-Path $dirs.Drives "MountPoints2.json")
  }
  finally {
    & reg.exe unload $tempHive2 | Out-Null
  }
}

# -------------------------
# FINAL: manifest + ZIP
# -------------------------
$manifest = [pscustomobject]@{
  CollectedAtUTC       = (Get-Date).ToUniversalTime()
  MachineName          = $env:COMPUTERNAME
  UserProfilePath      = $TargetProfilePath
  OutputRoot           = $OutputRoot
  Sections             = @("Printers","Bluetooth","WiFi(optional)","USB","Drives","RegistryExports","Logs")
  WifiXmlExported      = [bool]$ExportWifiXml
  WifiIncludedKeys     = [bool]$WifiIncludeKeys
}
$manifest | Safe-ExportJson -Path (Join-Path $OutputRoot "Manifest.json")

Write-Host "Creating ZIP archive..."
$zipPath = Join-Path (Split-Path $OutputRoot -Parent) ("$(Split-Path $OutputRoot -Leaf).zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $OutputRoot '*') -DestinationPath $zipPath -Force

Stop-Transcript | Out-Null

Write-Host "`nDone! Artifacts:"
Write-Host "  Root folder : $OutputRoot"
Write-Host "  ZIP archive : $zipPath"

