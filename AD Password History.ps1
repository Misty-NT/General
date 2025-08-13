# Get all enabled users, with the Last Password Set date
Get-ADUser -Filter * -Properties DisplayName, SamAccountName, PasswordLastSet |
    Select-Object DisplayName,
                  SamAccountName,
                  @{Name="PasswordLastChanged"; Expression={($_.PasswordLastSet).ToLocalTime()}} |
    Sort-Object PasswordLastChanged -Descending |
    Export-Csv -Path "C:\Temp\AD_UserPasswordLastChanged.csv" -NoTypeInformation

Write-Host "✅ Export complete. File saved to C:\NetworkTitan\AD_UserPasswordLastChanged.csv"
