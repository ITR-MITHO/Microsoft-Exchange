# Add FullAccess & Send-as permissions to the new sharedmailbox.
$CSV = Import-csv $home\dekstop\users.csv
ForEach ($C in $CSV)
{

$Name = $C.Name
Add-MailboxPermission -Identity SharedMailboxName -User $Name -AccessRights FullAccess
Get-Mailbox SharedMailboxName | Add-ADPermission -Extendedrights "Send-as" -User $Name

}

# List SMTP addresses for PublicFolder
Get-MailPublicfolder -Identity PublicFolderName | fl PrimarySMTPAddress, Emailaddresses

# Change PublicFolder SMTP Address
Set-MailPublicFolder -Identity PublicFolderName -PrimarySmtpAddress mail1000@domain.com -EmailAddressPolicyEnabled $false
Set-MailPublicFolder -Identity PublicFolderName -EmailAddresses mail1000@domain.com -EmailAddressPolicyEnabled $false

# Add SMTP addresses to the new shared mailbox
Set-Mailbox Vmail1 -PrimarySmtpAddress vmail@domain.com -EmailAddressPolicyEnabled $false
Set-Mailbox Vmail1 -EmailAddresses @{add="vmail@domain.com","Mail1@domain.com"}

# Export the public folder to a .PST-file via Outlook PST-export & save the file on your Exchange server
New-MailboxImportRequest -Mailbox SharedMailboxName -FilePath \\EXCH01\d$\PST-export\User.pst -Priority Highest -AcceptLargeDataLoss -BadItemLimit 500
