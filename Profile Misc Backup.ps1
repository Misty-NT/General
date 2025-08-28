# Wrapper: Export endpoint user config (printers, BT, Wi-Fi, USB, mapped drives)
# Runs as SYSTEM in Ninja, auto-detects target profile unless overridden.
# Params (as Ninja "Script Variables" or plain arguments):
#   TARGET_PROFILE  - e.g., C:\Users\scheduling  (default: auto)
#   OUTPUT_ROOT     - e.g., C:\NetworkTitan\Profile\EndpointBackup  (default: timestamped under C:\NetworkTitan\Profile)
#   EXPORT_WIFI     - true/false (default: false)
#   INCLUDE_KEYS    - true/false (default: false; only valid if EXPORT_WIFI=true)

param(
  [string]$TARGET_PROFILE = $env:TARGET_PROFILE,
  [string]$OUTPUT_ROOT    = $env:OUTPUT_ROOT,
  [string]$EXPORT_WIFI    = $env:EXPORT_WIFI,
  [string]$INCLUDE_KEYS   = $env:INCLUDE_KEYS
)

# Fall back to $Args[] if not provided as env vars
if (-not $TARGET_PROFILE -and $Args.Count -ge 1) { $TARGET_PROFILE = $Args[0] }
if (-not $OUTPUT_ROOT   -and $Args.Count -ge 2) { $OUTPUT_ROOT    = $Args[1] }
if (-not $EXPORT_WIFI   -and $Args.Count -ge 3) { $EXPORT_WIFI    = $Args[2] }
if (-not $INCLUDE_KEYS  -and $Args.Count -ge 4) { $INCLUDE_KEYS   = $Args[3] }

# Normalize booleans
function To-Bool($val, $default=$false) {
  if ($null -eq $val -or ($val -is [string] -and [string]::IsNullOrWhiteSpace($val))) { return $default }
  switch -Regex ($val.ToString()) {
    '^(1|true|yes|y)$'  { return $true }
    '^(0|false|no|n)$'  { return $false }
    default { return $default }
  }
}

$exportWifi  = To-Bool $EXPORT_WIFI $false
$includeKeys = To-Bool $INCLUDE_KEYS $false

# Auto-detect a reasonable profile if not supplied
function Get-ProfilePathAuto {
  try {
    # Prefer the *last interactive* user (not default/system)
    $owner = (Get-CimInstance Win32_ComputerSystem).Username  # DOMAIN\User or COMPUTER\User
    if ($owner) {
      $sam = ($owner -split '\\')[-1]
      $candidate = Join-Path 'C:\Users' $sam
      if (Test-Path $candidate) { return $candidate }
    }
  } catch {}
  # Fallback: newest non-default profile under C:\Users
  $fallback = Get-ChildItem 'C:\Users' -Directory |
    Where-Object { $_.Name -notin @('Public','Default','Default User','All Users','Administrator','Administrateur') } |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($fallback) { return $fallback.FullName }
  return $null
}

if (-not $TARGET_PROFILE -or -not (Test-Path $TARGET_PROFILE -PathType Container)) {
  $TARGET_PROFILE = Get-ProfilePathAuto
  if (-not $TARGET_PROFILE) {
    Write-Error "Could not determine a user profile path automatically. Supply TARGET_PROFILE."
    exit 1
  } else {
    Write-Host "Auto-selected profile: $TARGET_PROFILE"
  }
}

# Working paths
$baseDir   = 'C:\NetworkTitan\Profile'
$workDir   = Join-Path $baseDir 'Runner'
$scriptOut = Join-Path $workDir 'Export-EndpointUserConfig.ps1'

# Ensure folders
New-Item -ItemType Directory -Path $baseDir -Force | Out-Null
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

# --- Embed your child script (with transcript/zip order fixed) ---
$child = @'
<# Export-EndpointUserConfig.ps1 - patched to Stop-Transcript BEFORE zipping #>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidateScript({Test-Path $_ -PathType Container})]
  [string]$TargetProfilePath,

  [string]$OutputRoot,

  [switch]$ExportWifiXml,
  [switch]$WifiIncludeKeys
)

function New-OutDir { param([string]$Path) if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
function Safe-ExportJson {
  param([Parameter(ValueFromPipeline=$true)]$Object, [string]$Path)
  process { try { $Object | ConvertTo-Json -Depth 6 | Out-File -FilePath $Path -Encoding UTF8 } catch { ($_ | Out-String) | Out-File -FilePath $Path -Encoding UTF8 } }
}

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

# --- Printers (system-wide) ---
$printerList=@(); try {$printerList=Get-Printer -ErrorAction Stop} catch {Write-Warning "Get-Printer failed: $($_.Exception.Message)"}
$printerDrivers=@(); try {$printerDrivers=Get-PrinterDriver -ErrorAction Stop} catch {Write-Warning "Get-PrinterDriver failed: $($_.Exception.Message)"}
$printerPorts=@(); try {$printerPorts=Get-PrinterPort -ErrorAction Stop} catch {Write-Warning "Get-PrinterPort failed: $($_.Exception.Message)"}

$printConfigs = foreach ($p in $printerList) {
  try { $cfg=Get-PrintConfiguration -PrinterName $p.Name -ErrorAction Stop; [pscustomobject]@{PrinterName=$p.Name;Config=$cfg} }
  catch { [pscustomobject]@{PrinterName=$p.Name;Config=$null;Note="Get-PrintConfiguration failed: $($_.Exception.Message)"} }
}

$printerList    | Export-Csv (Join-Path $dirs.Printers "Printers.csv") -NoTypeInformation -Encoding UTF8
$printerDrivers | Export-Csv (Join-Path $dirs.Printers "PrinterDrivers.csv") -NoTypeInformation -Encoding UTF8
$printerPorts   | Export-Csv (Join-Path $dirs.Printers "PrinterPorts.csv") -NoTypeInformation -Encoding UTF8
$printConfigs   | Safe-ExportJson -Path (Join-Path $dirs.Printers "PrinterConfigs.json")

$defaultPrinterObj=$null; try {$defaultPrinterObj=Get-Printer | Where-Object {$_.Default -eq $true}} catch {}
$defaultPrinterWmi=$null; try {$defaultPrinterWmi=Get-CimInstance -Class Win32_Printer | Where-Object {$_.Default -eq $true}} catch {}
[pscustomobject]@{
  DefaultPrinterName_GetPrinter = $defaultPrinterObj.Name
  DefaultPrinterName_WMI        = $defaultPrinterWmi.Name
} | Safe-ExportJson -Path (Join-Path $dirs.Printers "DefaultPrinter_System.json")

# --- Per-user printers (load HKCU of target user) ---
$ntUserDat = Join-Path $TargetProfilePath "NTUSER.DAT"
if (-not (Test-Path $ntUserDat)) { Write-Warning "NTUSER.DAT not found at $ntUserDat. Skipping per-user registry exports." }
else {
  $tempHive = "HKU\TempUserHive_$stamp"
  & reg.exe load $tempHive $ntUserDat | Out-Null
  try {
    $userRegRoot="Registry::$tempHive"
    $userPrintersKey = Join-Path $userRegRoot "Software\Microsoft\Windows NT\CurrentVersion\Printers"
    $userWindowsKey  = Join-Path $userRegRoot "Software\Microsoft\Windows NT\CurrentVersion\Windows"

    & reg.exe export "$tempHive\Software\Microsoft\Windows NT\CurrentVersion\Printers" (Join-Path $dirs.RegExports "User_PrinterKeys.reg") /y | Out-Null
    & reg.exe export "$tempHive\Software\Microsoft\Windows NT\CurrentVersion\Windows"  (Join-Path $dirs.RegExports "User_WindowsKey.reg") /y | Out-Null

    $connections=$null;$devices=$null;$printerPorts=$null;$settings=$null
    try { $connections = Get-Item -Path (Join-Path $userPrintersKey "Connections") -ErrorAction Stop | Get-ChildItem -Recurse -ErrorAction SilentlyContinue } catch {}
    try { $devices     = Get-ItemProperty -Path (Join-Path $userWindowsKey  "Devices") -ErrorAction SilentlyContinue } catch {}
    try { $printerPorts= Get-ItemProperty -Path (Join-Path $userWindowsKey  "PrinterPorts") -ErrorAction SilentlyContinue } catch {}
    try {
      $settings = Get-ChildItem -Path (Join-Path $userPrintersKey "Settings") -ErrorAction SilentlyContinue | ForEach-Object {
        [pscustomobject]@{ Printer = $_.PSChildName; Values = (Get-ItemProperty -Path $_.PSPath) }
      }
    } catch {}

    $defaultPerUser=$null; try { $defaultPerUser = (Get-ItemProperty -Path $userWindowsKey -Name "Device" -ErrorAction Stop).Device } catch {}

    [pscustomobject]@{
      DefaultPrinter_PerUser = $defaultPerUser
      Connections_Subkeys    = $connections | Select-Object PSChildName, PSPath
      Devices_ValueBag       = $devices
      PrinterPorts_ValueBag  = $printerPorts
      Settings_PerPrinter    = $settings
    } | Safe-ExportJson -Path (Join-Path $dirs.Printers "PerUser_PrinterProfile.json")
  } finally {
    & reg.exe unload $tempHive | Out-Null
  }
}

# --- Bluetooth ---
$btOut=@()
try { $btPnp = Get-PnpDevice -Class Bluetooth -Status OK -ErrorAction Stop; $btOut += $btPnp | Select-Object FriendlyName, InstanceId, Present, Status, Problem } catch { Write-Warning "Get-PnpDevice (Bluetooth) failed: $($_.Exception.Message)" }
try {
  $bthKey = "HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Devices"
  if (Test-Path $bthKey) {
    $btReg = Get-ChildItem $bthKey -ErrorAction SilentlyContinue | ForEach-Object {
      $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
      [pscustomobject]@{ Source="Registry"; MacKey=$_.PSChildName; Properties=$props }
    }
    $btOut += $btReg
  }
} catch { Write-Warning "Bluetooth registry scrape failed: $($_.Exception.Message)" }
$btOut | Safe-ExportJson -Path (Join-Path $dirs.Bluetooth "Bluetooth_Inventory.json")

# --- Wi-Fi (optional) ---
if ($ExportWifiXml) {
  Write-Host "Exporting Wi-Fi profiles via netsh (XMLs)..."
  $opt = if ($WifiIncludeKeys) { " key=clear" } else { "" }
  $profilesTxt = Join-Path $dirs.WiFi "Profiles.txt"
  cmd.exe /c "netsh wlan show profiles > `"$profilesTxt`""
  cmd.exe /c "netsh wlan export profile folder=`"$($dirs.WiFi)`"$opt | more" | Out-Null
  if ($WifiIncludeKeys) {
@"
These Wi-Fi profile XMLs include CLEAR KEYS (passwords).
Treat as sensitive data and delete after use.
"@ | Out-File -FilePath (Join-Path $dirs.WiFi "WARNING.txt") -Encoding UTF8
  }
}

# --- USB history ---
$usbList=@()
try {
  $usbEnumKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR"
  if (Test-Path $usbEnumKey) {
    $usbList = Get-ChildItem $usbEnumKey -ErrorAction SilentlyContinue | ForEach-Object {
      $devName = $_.PSChildName
      Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue | ForEach-Object {
        $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{ Device=$devName; Instance=$_.PSChildName; Friendly=$props.FriendlyName; Mfg=$props.Mfg; ParentId=$props.ParentIdPrefix; Class=$props.Class }
      }
    }
  }
} catch { Write-Warning "USB history read failed: $($_.Exception.Message)" }
$usbList | Export-Csv (Join-Path $dirs.USB "USBSTOR.csv") -NoTypeInformation -Encoding UTF8

# --- Mapped drives ---
$currentMapped = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -like '\\*' } | Select-Object Name, Root, Used, Free, Description
$currentMapped | Export-Csv (Join-Path $dirs.Drives "MappedDrives_CurrentSession.csv") -NoTypeInformation -Encoding UTF8

if (Test-Path $ntUserDat) {
  $tempHive2 = "HKU\TempUserHive2_$stamp"
  & reg.exe load $tempHive2 $ntUserDat | Out-Null
  try {
    $root = "Registry::$tempHive2"
    $netKey = Join-Path $root "Network"
    $mpKey  = Join-Path $root "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
    $mappedPerUser=@(); $mountPoints=$null
    if (Test-Path $netKey) {
      $mappedPerUser += Get-ChildItem $netKey -ErrorAction SilentlyContinue | ForEach-Object {
        $p = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{ DriveLetter=$_.PSChildName; RemotePath=$p.RemotePath; UserName=$p.UserName; Provider=$p.ProviderName; DeferFlags=$p.DeferFlags }
      }
    }
    if (Test-Path $mpKey) { $mountPoints = Get-ChildItem $mpKey -ErrorAction SilentlyContinue | Select-Object PSChildName, PSPath }
    $mappedPerUser | Export-Csv (Join-Path $dirs.Drives "MappedDrives_PerUser.csv") -NoTypeInformation -Encoding UTF8
    $mountPoints  | Safe-ExportJson -Path (Join-Path $dirs.Drives "MountPoints2.json")
  } finally { & reg.exe unload $tempHive2 | Out-Null }
}

# --- Manifest ---
$manifest = [pscustomobject]@{
  CollectedAtUTC   = (Get-Date).ToUniversalTime()
  MachineName      = $env:COMPUTERNAME
  UserProfilePath  = $TargetProfilePath
  OutputRoot       = $OutputRoot
  Sections         = @("Printers","Bluetooth","WiFi(optional)","USB","Drives","RegistryExports","Logs")
  WifiXmlExported  = [bool]$ExportWifiXml
  WifiIncludedKeys = [bool]$WifiIncludeKeys
}
$manifest | Safe-ExportJson -Path (Join-Path $OutputRoot "Manifest.json")

# --- IMPORTANT FIX: Close transcript before zipping to avoid file lock ---
Stop-Transcript | Out-Null

# --- ZIP (now safe) ---
$zipPath = Join-Path (Split-Path $OutputRoot -Parent) ("$(Split-Path $OutputRoot -Leaf).zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $OutputRoot '*') -DestinationPath $zipPath -Force

Write-Host "`nDone! Artifacts:"
Write-Host "  Root folder : $OutputRoot"
Write-Host "  ZIP archive : $zipPath"
'@

# Write child script to disk
$child | Out-File -FilePath $scriptOut -Encoding UTF8 -Force

# Build args for child
$argList = @('-ExecutionPolicy','Bypass','-File', $scriptOut, '-TargetProfilePath', $TARGET_PROFILE)
if ($OUTPUT_ROOT)  { $argList += @('-OutputRoot', $OUTPUT_ROOT) }
if ($exportWifi)   { $argList += '-ExportWifiXml' }
if ($includeKeys)  { $argList += '-WifiIncludeKeys' }

Write-Host "Running backup with:"
Write-Host "  TargetProfilePath = $TARGET_PROFILE"
if ($OUTPUT_ROOT)  { Write-Host "  OutputRoot       = $OUTPUT_ROOT" }
Write-Host "  ExportWifiXml     = $exportWifi"
Write-Host "  WifiIncludeKeys   = $includeKeys"

# Invoke child
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName               = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$psi.Arguments              = ($argList -join ' ')
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError  = $true
$psi.UseShellExecute        = $false
$proc = [System.Diagnostics.Process]::Start($psi)
$stdout = $proc.StandardOutput.ReadToEnd()
$stderr = $proc.StandardError.ReadToEnd()
$proc.WaitForExit()

# Echo outputs to Ninja Activity Log
if ($stdout) { Write-Host $stdout }
if ($stderr) { Write-Warning $stderr }

# Optional: write a small pointer file for techs
try {
  $hint = Join-Path $workDir "last_run.txt"
  @(
    "Ran: $(Get-Date -Format s)"
    "TargetProfilePath: $TARGET_PROFILE"
    "OutputRoot: $OUTPUT_ROOT"
    "ExportWifiXml: $exportWifi"
    "WifiIncludeKeys: $includeKeys"
  ) | Out-File -FilePath $hint -Encoding UTF8 -Force
} catch {}

exit $proc.ExitCode
