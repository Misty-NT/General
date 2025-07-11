# Huntress SIEM Local Audit Policy Setup
try {
    Write-Output "Starting audit policy and log config..."

    # Apply audit policies
    $AuditCmds = @(
        'auditpol /set /subcategory:"{0CCE923F-69AE-11D9-BED3-505054503030}","{0CCE9242-69AE-11D9-BED3-505054503030}","{0CCE9240-69AE-11D9-BED3-505054503030}","{0CCE9236-69AE-11D9-BED3-505054503030}","{0CCE9238-69AE-11D9-BED3-505054503030}","{0CCE9237-69AE-11D9-BED3-505054503030}","{0CCE9235-69AE-11D9-BED3-505054503030}","{0CCE923B-69AE-11D9-BED3-505054503030}","{0CCE9215-69AE-11D9-BED3-505054503030}","{0CCE9243-69AE-11D9-BED3-505054503030}","{0CCE921C-69AE-11D9-BED3-505054503030}","{0CCE9244-69AE-11D9-BED3-505054503030}","{0CCE9224-69AE-11D9-BED3-505054503030}","{0CCE921F-69AE-11D9-BED3-505054503030}","{0CCE9227-69AE-11D9-BED3-505054503030}","{0CCE9245-69AE-11D9-BED3-505054503030}","{0CCE9232-69AE-11D9-BED3-505054503030}","{0CCE9234-69AE-11D9-BED3-505054503030}","{0CCE9228-69AE-11D9-BED3-505054503030}","{0CCE9214-69AE-11D9-BED3-505054503030}","{0CCE9212-69AE-11D9-BED3-505054503030}" /success:enable /failure:enable',
        'auditpol /set /subcategory:"{0CCE923A-69AE-11D9-BED3-505054503030}","{0CCE9248-69AE-11D9-BED3-505054503030}","{0CCE923C-69AE-11D9-BED3-505054503030}","{0CCE9216-69AE-11D9-BED3-505054503030}","{0CCE921B-69AE-11D9-BED3-505054503030}","{0CCE922F-69AE-11D9-BED3-505054503030}","{0CCE9230-69AE-11D9-BED3-505054503030}","{0CCE9231-69AE-11D9-BED3-505054503030}","{0CCE9233-69AE-11D9-BED3-505054503030}","{0CCE9210-69AE-11D9-BED3-505054503030}","{0CCE9211-69AE-11D9-BED3-505054503030}" /success:enable',
        'auditpol /set /subcategory:"{0CCE9217-69AE-11D9-BED3-505054503030}","{0CCE9226-69AE-11D9-BED3-505054503030}" /failure:enable'
    )
    foreach ($cmd in $AuditCmds) {
        Write-Output "Running: $cmd"
        Invoke-Expression $cmd
    }

    # PowerShell Logging
    $base = "HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\PowerShell"
    New-Item -Path "$base\ScriptBlockLogging" -Force | Out-Null
    Set-ItemProperty -Path "$base\ScriptBlockLogging" -Name EnableScriptBlockLogging -Value 1
    New-Item -Path "$base\ModuleLogging" -Force | Out-Null
    Set-ItemProperty -Path "$base\ModuleLogging" -Name EnableModuleLogging -Value 1
    New-Item -Path "$base\ModuleLogging\ModuleNames" -Force | Out-Null
    New-ItemProperty -Path "$base\ModuleLogging\ModuleNames" -Name "*" -Value "*" -Force | Out-Null

    # Event Log Size & Retention
    Limit-EventLog -LogName Security -MaximumSize 512000KB -OverflowAction OverwriteAsNeeded

    # Disable noisy system object auditing
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "AuditBaseObjects" -Value 0

    Write-Output "Audit policy configuration complete."
}
catch {
    Write-Output "Error during configuration: $_"
}
