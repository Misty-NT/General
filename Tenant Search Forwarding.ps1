Get-Mailbox -ResultSize Unlimited |
Where-Object {
    $_.ForwardingSMTPAddress -ne $null -or
    $_.ForwardingAddress -ne $null
} |
Select-Object DisplayName, UserPrincipalName, ForwardingSMTPAddress, ForwardingAddress, DeliverToMailboxAndForward |
Export-Csv -Path "c:\NT\MailboxForwardingReport.csv" -NoTypeInformation -Encoding UTF8

# This will loop through and export the emails being forwarded into a csv to the path. 
