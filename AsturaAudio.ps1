# Restore-AudioDefaults.ps1
# Reads C:\NetworkTitan\audio-defaults.json and forces those devices as defaults
# (Playback/Recording for both Multimedia and Communications roles)

$ErrorActionPreference = 'Stop'

$configFile = 'C:\NetworkTitan\audio-defaults.json'
if (-not (Test-Path $configFile)) {
  Write-Error "Config not found: $configFile. Run Export-AudioDefaults.ps1 first."
  exit 1
}

# Ensure module
$moduleName = 'AudioDeviceCmdlets'
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
  try {
    Install-Module -Name $moduleName -Scope CurrentUser -Force -ErrorAction Stop
  } catch {
    Write-Error "Failed to install $moduleName. Run PowerShell as the user and allow: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
    exit 1
  }
}
Import-Module $moduleName -ErrorAction Stop

# Load config
try {
  $cfg = Get-Content $configFile -Raw | ConvertFrom-Json
} catch {
  Write-Error "Could not read or parse $configFile"
  exit 1
}

# Helper: resolve device by Id first, then name (exact, then fuzzy)
function Resolve-AudioDevice {
  param(
    [string]$Type,       # 'Playback' or 'Recording'
    [string]$Id,
    [string]$Name
  )
  $list = Get-AudioDevice -List | Where-Object { $_.Type -eq $Type }

  if ($Id) {
    $byId = $list | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
    if ($byId) { return $byId }
  }
  if ($Name) {
    $exact = $list | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if ($exact) { return $exact }
    $fuzzy = $list | Where-Object { $_.Name -like "*$Name*" } | Select-Object -First 1
    if ($fuzzy) { return $fuzzy }
  }
  return $null
}

$targets = @(
  @{ Role = 'Multimedia';    Type = 'Playback';  Source = $cfg.Playback        },
  @{ Role = 'Communications';Type = 'Playback';  Source = $cfg.PlaybackComms   },
  @{ Role = 'Multimedia';    Type = 'Recording'; Source = $cfg.Recording       },
  @{ Role = 'Communications';Type = 'Recording'; Source = $cfg.RecordingComms  }
)

$allOk = $true
foreach ($t in $targets) {
  $dev = Resolve-AudioDevice -Type $t.Type -Id $t.Source.Id -Name $t.Source.Name
  if (-not $dev) {
    Write-Warning "Could not find $($t.Type) device for saved '$($t.Source.Name)'. Ensure it is connected."
    $allOk = $false
    continue
  }
  try {
    Set-DefaultAudioDevice -Id $dev.Id -Role $t.Role | Out-Null
    Write-Host "Set $($t.Type) -> '$($dev.Name)' as default ($($t.Role))." -ForegroundColor Cyan
  } catch {
    Write-Warning "Failed to set $($t.Type) '$($dev.Name)' ($($t.Role)): $($_.Exception.Message)"
    $allOk = $false
  }
}

if ($allOk) {
  Write-Host "Audio defaults restored successfully." -ForegroundColor Green
} else {
  Write-Host "One or more roles could not be set. Plug in devices and run again." -ForegroundColor Yellow
}
