# Mailbox Size Total - 

$mailboxes = @(
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,",
    "email@domain.com,"
)

foreach ($email in $mailboxes) {
    $params = [ordered]@{
        Identity = $email
    }

    Get-MailboxStatistics @params | Select-Object DisplayName, TotalItemSize
}


***********************************************************************************
===================================================================================

# Misc Bloat Steps 
# Disable Registry & task to prevent reinstalling windows bloat
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent" /v DisableConsumerFeatures /t REG_DWORD /d 1 /f

New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Force

New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableConsumerFeatures" -Value 1 -PropertyType DWord -Force


reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent" /v DisableConsumerFeatures /t REG_DWORD /d 1 /f

schtasks /Change /TN "Microsoft\Windows\CloudExperienceHost\CreateObjectTask" /Disable
schtasks /Change /TN "Microsoft\Windows\Consumer Experiences\CleanUpTemporaryState" /Disable
schtasks /Change /TN "Microsoft\Windows\Consumer Experiences\StartupAppTask" /Disable

***********************************************************************************
===================================================================================
# Commands to clean temp and etc
del /s /q "%localappdata%\Temp\*"
del /s /q ""
del /s /q "C:\Windows\Temp\*"

# Clear Windows Update Downloads
net stop wuauserv
net stop bits
rd /s /q "C:\Windows\SoftwareDistribution\Download"
net start wuauserv
net start bits

# Check Shadow Copy Size for All Drives
vssadmin list shadowstorage

#To Check for a Specific Drive

vssadmin list shadowstorage /for=C:
# CMD to Empty All Recycle Bins
rd /s /q C:\$Recycle.Bin
# This affects all drives — repeat for other drives if needed:
rd /s /q D:\$Recycle.Bin
rd /s /q E:\$Recycle.Bin
# To run for all drives in one line (example for C, D, E):
rd /s /q C:\$Recycle.Bin & rd /s /q D:\$Recycle.Bin & rd /s /q E:\$Recycle.Bin

# Windows Update Cache
net stop wuauserv
net stop bits
rd /s /q C:\Windows\SoftwareDistribution\Download
net start wuauserv
net start bits

# Recycle Bin
rd /s /q C:\$Recycle.Bin

# Windows.old Folder (after major updates)
rd /s /q C:\Windows.old
#  Prefetch Files
del /s /q "C:\Users\robert\AppData\Roaming\*"
del /s /q  "C:\NetworkTitan\WICleanup\*"
del /s /q  "C:\Recovery\*"
del /s /q   "C:\OneDriveTemp\S-1-12-1-3483278516-1171702208-4040308392-1566361512\*"
del /s /q  "C:\Users\jack.sighton\AppData\Local\Google\Chrome\User Data\Default\Service Worker\CacheStorage\*"
del /s /q C:\Windows\System32\config\systemprofile\AppData\Local\CrashDumps\*
#  Downloaded Program Files
C:\Windows\System32\config\systemprofile\AppData\Local\Microsoft\*
del /s /q  C:\Windows\System32\config\systemprofile\AppData\Local\Microsoft\Office\*
del /s /q  C:\Windows\System32\config\systemprofile\AppData\Local\Microsoft\Office\16.0\WebServiceCache\*
rd /s /q "C:\Windows\Downloaded Program Files"
#  Log Files
del /f /s /q C:\Windows\Logs\*.* >nul 2>&1
dir /s /a "C:\Windows\Installer"
dir /s /a "C:\System Volume Information"
dir /s /a "C:\Users\localadmin"

***********************************************************************************
===================================================================================
takeown /f  "C:\Windows.old\Users\All Users\Dell" /r /d y

del /s /q "C:\Windows.old\Users\All Users\Dell"
takeown /f "C:\Windows.old\Users\bfajardo" /r /d y
icacls "C:\Windows.old\Users\All Users\Dell" /grant %username%:F /t
icacls "C:\Windows.old\Users\bfajardo" /grant %username%:F /t

# Take Ownership
del /s /q  "C:\ProgramData\Dell"
takeown /f "C:\Windows.old" /r /d y
icacls "C:\Windows.old" /grant "%USERNAME%:F" /t /c
attrib -r -s -h "C:\Windows.old" /s /d

***********************************************************************************
===================================================================================

# Drill into a Parent Folder to get the size of the contents 
$targetPath = "C:\Path\To\You\Location"  # Automatically sets to user's roaming folder

# Get top-level folders and calculate sizes
Get-ChildItem -Path $targetPath -Directory | ForEach-Object {
    $folder = $_.FullName
    $size = (Get-ChildItem -Path $folder -Recurse -ErrorAction SilentlyContinue | 
             Measure-Object -Property Length -Sum).Sum
    [PSCustomObject]@{
        Folder = $_.Name
        SizeGB = '{0:N2}' -f ($size / 1GB)
    }
} | Sort-Object -Property SizeGB -Descending | Format-Table -AutoSize

***********************************************************************************
===================================================================================
# Purview Search 
# 2. Define search name (must be unique)
$searchName = "Search-Phish2025-17-6-" + (Get-Date -Format "yyyyMMdd-HHmm")

# 3. Build the KQL query
# Sent today = sent>=yyyy-MM-dd AND sent<=(yyyy-MM-ddT23:59:59)
$today = (Get-Date).ToString("yyyy-MM-dd")
$tomorrow = (Get-Date).AddDays(1).ToString("yyyy-MM-dd")

$query = 'from:"brett@advisorlogistics.com" AND subject:"Brett Wheeler, CFP - Founder*" AND (sent>=' + $today + ' AND sent<' + $tomorrow + ')'

# 4. Create the compliance search (all mailboxes by default)
New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $query

# 5. Start the search
Start-ComplianceSearch -Identity $searchName

# Optional: check status
Get-ComplianceSearch -Identity $searchName
