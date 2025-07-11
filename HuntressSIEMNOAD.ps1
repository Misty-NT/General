# =============================
# Huntress SIEM Local Log Setup
# =============================

Write-Host "Starting local Windows audit policy and log configuration..." -ForegroundColor Cyan

# --- 1. Set Advanced Audit Policies (Success + Failure) ---
$AuditCmds = @(
    'auditpol /set /subcategory:"{0CCE923F-69AE-11D9-BED3-505054503030}","{0CCE9242-69AE-11D9-BED3-505054503030}","{0CCE9240-69AE-11D9-BED3-505054503030}","{0CCE9236-69AE-11D9-BED3-505054503030}","{0CCE9238-69AE-11D9-BED3-505054503030}","{0CCE9237-69AE-11D9-BED3-505054503030}","{0CCE9235-69AE-11D9-BED3-505054503030}","{0CCE923B-69AE-11D9-BED3-505054503030}","{0CCE9215-69AE-11D9-BED3-505054503030}","{0CCE9243-69AE-11D9-BED3-505054503030}","{0CCE921C-69AE-11D9-BED3-505054503030}","{0CCE9244-69AE-11D9-BED3-505054503030}","{0CCE9224-69AE-11D9-BED3-505054503030}","{0CCE921F-69AE-11D9-BED3-505054503030}","{0CCE9227-69AE-11D9-BED3-505054503030}","{0CCE9245-69AE-11D9-BED3-505054503030}","{0CCE9232-69AE-11D9-BED3-505054503030}","{0CCE9234-69AE-11D9-BED3-505054503030}","{0CCE9228-69AE-11D9-BED3-505054503030}","{0CCE9214-69AE-11D9-BED3-505054503030}","{0CCE9212-69AE-11D9-BED3-505054503030}" /success:enable /failure:enable',
    'auditpol /set /subcategory:"{0CCE923A-69AE-11D9-BED3-505054503030}","{0CCE9248-69AE-11D9-BED3-505054503030}","{0CCE923C-69AE-11D9-BED3-505054503030}","{0CCE9216-69AE-11D9-BED3-505054503030}","{0CCE921B-69AE-11D9-BED3-505054503030}","{0CCE922F-69AE-11D9-BED3-505054503030}","{0CCE9230-69AE-11D9-BED3-505054503030}","{0CCE9231-69AE-11D9-BED3-505054503030}","{0CCE9233-69AE-11D9-BED3-505054503030}","{0CCE9210-69AE-11D9-BED3-505054503030}","{0CCE9211-69AE-11D9-BED3-505054503030}" /success:enable',
    'auditpol /set /subcategory:"{0CCE9217-69AE-11D9-BED3-505054503030}","{0CCE9226-69AE-11D9-BED3-505054503030}" /failure:enable'
)

foreach ($cmd in $AuditCmds) {
    Write-Host "Applying: $cmd" -ForegroundColor DarkYellow
    Invoke-Expression $cmd
}

# --- 2. Configure PowerShell ScriptBlock and Module Logging ---
Write-Host "Configuring PowerShell script and module logging..." -ForegroundColor Cyan

$psBase = "HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\PowerShell"
New-Item -Path "$psBase\ScriptBlockLogging" -Force | Out-Null
Set-ItemProperty -Path "$psBase\ScriptBlockLogging" -Name EnableScriptBlockLogging -Value 1

New-Item -Path "$psBase\ModuleLogging" -Force | Out-Null
Set-ItemProperty -Path "$psBase\ModuleLogging" -Name EnableModuleLogging -Value 1
New-Item -Path "$psBase\ModuleLogging\ModuleNames" -Force | Out-Null
New-ItemProperty -Path "$psBase\ModuleLogging\ModuleNames" -Name "*" -Value "*" -Force | Out-Null

# --- 3. Configure Event Log Size and Retention ---
Write-Host "Setting Security Event Log max size to 512MB and retention to OverwriteAsNeeded..." -ForegroundColor Cyan
Limit-EventLog -LogName Security -MaximumSize 512000KB -OverflowAction OverwriteAsNeeded

# --- 4. Optional: Disable 'Audit the access of global system objects' ---
Write-Host "Disabling global system object auditing to prevent log spam..." -ForegroundColor Cyan
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "AuditBaseObjects" -Value 0

# --- 5. Done ---
Write-Host "Local audit policy and event log settings configured!" -ForegroundColor Green
