# Export-EndpointUserConfig.ps1
# Minimal clean ASCII version

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidateScript({Test-Path $_ -PathType Container})]
  [string]$TargetProfilePath,
  [string]$OutputRoot,
  [switch]$ExportWifiXml,
  [switch]$WifiIncludeKeys
)

function New-OutDir($p){ if(-not (Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } }
function ToJsonFile($o,$p){ try{$o|ConvertTo-Json -Depth 6|Out-File -FilePath $p -Encoding UTF8}catch{($_|Out-String)|Out-File -FilePath $p -Encoding UTF8} }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
if(-not $OutputRoot){ $OutputRoot = "C:\Temp\EndpointBackup_$stamp" }
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

Write-Host "Target profile: $TargetProfilePath"
Write-Host "Output folder : $OutputRoot"

# ---------------- PRINTERS (system-wide) ----------------
Write-Host "Collecting printers..."
$printerList = @(); $printerDrivers = @(); $printerPorts = @()

try { $printerList   = Get-Printer -ErrorAction Stop } catch {}
try { $printerDrivers= Get-PrinterDriver -ErrorAction Stop } catch {}
try { $printerPorts  = Get-PrinterPort -ErrorAction Stop } catch {}

$printConfigs = foreach($p in $printerList){
  try{
    $cfg = Get-PrintConfiguration -PrinterName $p.Name -ErrorAction Stop
    [pscustomobject]@{PrinterName=$p.Name;Config=$cfg}
  }catch{
    [pscustomobject]@{PrinterName=$p.Name;Config=$null;Note="Get-PrintConfiguration failed"}
  }
}

$printerList    | Export-Csv (Join-Path $dirs.Printers "Printers.csv") -NoTypeInformation -Encoding UTF8
$printerDrivers | Export-Csv (Join-Path $dirs.Printers "PrinterDrivers.csv") -NoTypeInformation -Encoding UTF8
$printerPorts   | Export-Csv (Join-Path $dirs.Printers "PrinterPorts.csv") -NoTypeInformation -Encoding UTF8
ToJsonFile $printConfigs (Join-Path $dirs.Printers "PrinterConfigs.json")

$def1=$null;$def2=$null
try{$def1=(Get-Printer|Where-Object{$_.Default}).Name}catch{}
try{$def2=(Get-CimInstance Win32_Printer|Where-Object{$_.Default}).Name}catch{}
[pscustomobject]@{DefaultPrinter_GetPrinter=$def1;DefaultPrinter_WMI=$def2} |
  ToJsonFile -p (Join-Path $dirs.Printers "DefaultPrinter_System.json")

# ------------- PER-USER PRINTERS (load HKCU of target) -------------
Write-Host "Loading user hive for per-user printer info..."
$ntUserDat = Join-Path $TargetProfilePath "NTUSER.DAT"
if(Test-Path $ntUserDat){
  $tempHive = "HKU\TempUser_$stamp"
  & reg.exe load $tempHive $ntUserDat | Out-Null
  try{
    $root = "Registry::$tempHive"
    $kPrinters = Join-Path $root "Software\Microsoft\Windows NT\CurrentVersion\Printers"
    $kWindows  = Join-Path $root "Software\Microsoft\Windows NT\CurrentVersion\Windows"

    & reg.exe export "$tempHive\Software\Microsoft\Windows NT\CurrentVersion\Printers" (Join-Path $dirs.RegExports "User_PrinterKeys.reg") /y | Out-Null
    & reg.exe export "$tempHive\Software\Microsoft\Windows NT\CurrentVersion\Windows"  (Join-Path $dirs.RegExports "User_WindowsKey.reg") /y | Out-Null

    $connections=$null;$devices=$null;$printerPortsUser=$null;$settings=$null;$defaultPerUser=$null
    try{$connections=Get-Item -Path (Join-Path $kPrinters "Connections") -ErrorAction Stop | Get-ChildItem -Recurse -ErrorAction SilentlyContinue}catch{}
    try{$devices=Get-ItemProperty -Path (Join-Path $kWindows "Devices") -ErrorAction SilentlyContinue}catch{}
    try{$printerPortsUser=Get-ItemProperty -Path (Join-Path $kWindows "PrinterPorts") -ErrorAction SilentlyContinue}catch{}
    try{$settings=Get-ChildItem -Path (Join-Path $kPrinters "Settings") -ErrorAction SilentlyContinue | ForEach-Object{
          [pscustomobject]@{Printer=$_.PSChildName;Values=(Get-ItemProperty -Path $_.PSPath)}
        }}catch{}
    try{$defaultPerUser=(Get-ItemProperty -Path $kWindows -Name Device -ErrorAction Stop).Device}catch{}

    [pscustomobject]@{
      DefaultPrinter_PerUser = $defaultPerUser
      Connections_Subkeys    = $connections | Select-Object PSChildName,PSPath
      Devices_ValueBag       = $devices
      PrinterPorts_ValueBag  = $printerPortsUser
      Settings_PerPrinter    = $settings
    } | ToJsonFile -p (Join-Path $dirs.Printers "PerUser_PrinterProfile.json")
  } finally {
    & reg.exe unload $tempHive | Out-Null
  }
}else{
  Write-Warning "NTUSER.DAT not found at $ntUserDat. Skipping per-user printer capture."
}

# ---------------- BLUETOOTH ----------------
Write-Host "Collecting Bluetooth..."
$btOut=@()
try{
  $btPnp = Get-PnpDevice -Class Bluetooth -Status OK -ErrorAction Stop
  $btOut += $btPnp | Select-Object FriendlyName,InstanceId,Present,Status,Problem
}catch{}
try{
  $bthKey = "HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Devices"
  if(Test-Path $bthKey){
    $btOut += Get-ChildItem $bthKey | ForEach-Object{
      [pscustomobject]@{Source="Registry";MacKey=$_.PSChildName;Properties=(Get-ItemProperty -Path $_.PSPath)}
    }
  }
}catch{}
ToJsonFile $btOut (Join-Path $dirs.Bluetooth "Bluetooth_Inventory.json")

# ---------------- WIFI (optional) ----------------
if ($ExportWifiXml) {
  Write-Host "Exporting Wi-Fi profiles..."
  $profilesTxt = Join-Path $dirs.WiFi "Profiles.txt"
  cmd.exe /c "netsh wlan show profiles > `"$profilesTxt`""

  $opt = ""
  if ($WifiIncludeKeys) { $opt = " key=clear" }

  cmd.exe /c "netsh wlan export profile folder=`"$($dirs.WiFi)`"$opt | more" | Out-Null

  if ($WifiIncludeKeys) {
    $warning = "These Wi-Fi XMLs include clear keys (passwords). Protect or delete after use.`r`n"
    Set-Content -Path (Join-Path $dirs.WiFi 'WARNING.txt') -Value $warning -Encoding UTF8
  }
}
# ---------------- USB STORAGE HISTORY ----------------
Write-Host "Collecting USB storage history..."
$usbList=@()
try{
  $usbEnum = "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR"
  if(Test-Path $usbEnum){
    $usbList = Get-ChildItem $usbEnum -ErrorAction SilentlyContinue | ForEach-Object{
      $dev=$_.PSChildName
      Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue | ForEach-Object{
        $p=Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{Device=$dev;Instance=$_.PSChildName;Friendly=$p.FriendlyName;Mfg=$p.Mfg;ParentId=$p.ParentIdPrefix;Class=$p.Class}
      }
    }
  }
}catch{
  Write-Warning "USB read failed: $($_.Exception.Message)"
}
$usbList | Export-Csv (Join-Path $dirs.USB "USBSTOR.csv") -NoTypeInformation -Encoding UTF8

# ---------------- MAPPED DRIVES ----------------
Write-Host "Collecting mapped drives..."
$curr = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -like '\\*' } |
        Select-Object Name,Root,Used,Free,Description
$curr | Export-Csv (Join-Path $dirs.Drives "MappedDrives_CurrentSession.csv") -NoTypeInformation -Encoding UTF8

if(Test-Path $ntUserDat){
  $temp2 = "HKU\TempUser2_$stamp"
  & reg.exe load $temp2 $ntUserDat | Out-Null
  try{
    $root = "Registry::$temp2"
    $netKey = Join-Path $root "Network"
    $mpKey  = Join-Path $root "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
    $mapped=@()
    if(Test-Path $netKey){
      $mapped = Get-ChildItem $netKey -ErrorAction SilentlyContinue | ForEach-Object{
        $p=Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{DriveLetter=$_.PSChildName;RemotePath=$p.RemotePath;UserName=$p.UserName;Provider=$p.ProviderName;DeferFlags=$p.DeferFlags}
      }
    }
    $mounts=$null
    if(Test-Path $mpKey){ $mounts = Get-ChildItem $mpKey -ErrorAction SilentlyContinue | Select-Object PSChildName,PSPath }
    $mapped | Export-Csv (Join-Path $dirs.Drives "MappedDrives_PerUser.csv") -NoTypeInformation -Encoding UTF8
    ToJsonFile $mounts (Join-Path $dirs.Drives "MountPoints2.json")
  } finally {
    & reg.exe unload $temp2 | Out-Null
  }
}

# ---------------- MANIFEST + ZIP ----------------
$manifest = [pscustomobject]@{
  CollectedAtUTC = (Get-Date).ToUniversalTime()
  ComputerName   = $env:COMPUTERNAME
  UserProfile    = $TargetProfilePath
  OutputRoot     = $OutputRoot
}
ToJsonFile $manifest (Join-Path $OutputRoot "Manifest.json")

$zipPath = Join-Path (Split-Path $OutputRoot -Parent) ("$(Split-Path $OutputRoot -Leaf).zip")
if(Test-Path $zipPath){ Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $OutputRoot '*') -DestinationPath $zipPath -Force

Stop-Transcript | Out-Null
Write-Host "Done. Output: $OutputRoot"
Write-Host "ZIP: $zipPath"


