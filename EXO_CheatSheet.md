# Helpful links

Exchange Online Administrative Center  
https://admin.exchange.microsoft.com  

**Exchange Online Protection**  
Policies and Rules: https://security.microsoft.com/threatpolicy  

Explorer: https://security.microsoft.com/threatexplorer  

Legacy Retention Policies: https://purview.microsoft.com/datalifecyclemanagement/exchange/retentionpolicies  

Quarantine: https://security.microsoft.com/quarantine  
E-mail that sends quarantine notifications: Quarantine@messaging.microsoft.com  

**Exchange Online URLs and IPs**  
https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide#exchange-online  

**Inbound DANE**  
https://techcommunity.microsoft.com/blog/exchange/implementing-inbound-smtp-dane-with-dnssec-for-exchange-online-mail-flow/3939694  

**Exchange Online Limits**  
https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits  

**M365 Licenses**  
https://m365maps.com/

# Mailflow
**SPF:** v=spf1 spf.protection.outlook.com -all  
**Random SPF:** v=spf1 redirect=domain.com  
**DMARC**: v=DMARC1; p=reject; pct=100; adkim=s; aspf=s  
**MX:** domain-com.mail.protection.outlook.com  
**MX-DANE:** domain-com.l-v1.mx.microsoft

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

**10 days messagetrace**  
Get-MessageTraceV2 -ResultSize 5000 -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date)    

# Outlook New Registry Keys
**Hide Outlook New Button**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: HideNewOutlookToggle  
Value: 1


**Stop Outlook from becoming Outlook NEW**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: DoNewOutlookAutoMigration  
Value: 0  
