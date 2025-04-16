# Helpful links

Exchange Online Administrative Center  
https://admin.exchange.microsoft.com  

Exchange Online Protection Settings  
https://security.microsoft.com/threatpolicy  

Quarantine  
https://security.microsoft.com/quarantine  

Exchange Online URLs and IPs  
https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide#exchange-online  

Inbound DANE  
https://techcommunity.microsoft.com/blog/exchange/implementing-inbound-smtp-dane-with-dnssec-for-exchange-online-mail-flow/3939694  

Legacy Retention Policies  
https://purview.microsoft.com/datalifecyclemanagement/exchange/retentionpolicies  

Exchange Online Limits 
https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits  

# Move requests
Get-Moverequest | Get-MoverequestStatistics  
Set-MoveRequest MAILBOX -SkippedItemApprovalTime $(Get-Date).ToUniversalTime()  
Set-Moverequest MAILBOX -Completeafter 1  

# Permissions
Set-MailboxFolderPermission ALIAS:\Calendar -User Default -AccessRights LimitedDetails  
Add-RecipientPermission MAILBOX -Trustee USERNAME -AccessRights Sendas -Confirm:$false  
Add-MailboxPermission MAILBOX -User USERNAME -AccessRights FullAccess -Automapping $true  

# Nice to know
Set-User MAILBOX -PermanentlyClearPreviousMailboxInfo -confirm:$false  
Set-Mailbox MAILBOX -ApplyMandatoryProperties  
Get-MessageTrace -Start (Get-date).AddDays(-10) -End (Get-Date)  

# Mailflow
**SPF:** spf.protection.outlook.com  
**MX:** domain-dk.mail.protection.outlook.com  
**MX-DANE:** domain-dk.l-v1.mx.microsoft  

# Outlook New
**Hide Outlook New Button**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: HideNewOutlookToggle  
Value: 00000000  


**Stop automatic Outlook New Upgrade**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: DoNewOutlookAutoMigration  
Value: 00000000  
