# Lists all accounts that do NOT require Kerberos pre-auth
Get-ADUser -Filter * -Properties DoesNotRequirePreAuth |
Where-Object {$_.DoesNotRequirePreAuth -eq $true} |
Select-Object SamAccountName, UserPrincipalName
