# Helpful links

Exchange Online Administrative Center  
https://admin.exchange.microsoft.com  

Exchange Online Protection Settings & Recommendations  
https://security.microsoft.com/threatpolicy  
https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365  

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

M365 Licenses  
https://m365maps.com/

# PowerShell Commands
**Nice to know commands**  
Set-User MAILBOX -PermanentlyClearPreviousMailboxInfo -confirm:$false  
Set-Mailbox MAILBOX -ApplyMandatoryProperties  

**Show moverequest status in percentage**  
Get-Moverequest | Get-MoverequestStatistics  

**Approve migration with baditems immediatly**  
Set-MoveRequest MAILBOX -SkippedItemApprovalTime $(Get-Date).ToUniversalTime()  

**Complete moverequest immediatly**  
Set-Moverequest MAILBOX -Completeafter 1  

**Commands to set permissions**  
Set-MailboxFolderPermission ALIAS:\Calendar -User Default -AccessRights LimitedDetails  
Add-RecipientPermission MAILBOX -Trustee USERNAME -AccessRights Sendas -Confirm:$false  
Add-MailboxPermission MAILBOX -User USERNAME -AccessRights FullAccess -Automapping $true  

# Mailflow
**SPF:** spf.protection.outlook.com  
**MX:** domain-dk.mail.protection.outlook.com  
**MX-DANE:** domain-dk.l-v1.mx.microsoft  

# Outlook New Registries
**Hide Outlook New Button**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: HideNewOutlookToggle  
Value: 1


**Stop automatic Outlook New Upgrade**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: DoNewOutlookAutoMigration  
Value: 0  
