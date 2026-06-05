# Helpful links

**Exchange Online Administrative Center**  
https://admin.exchange.microsoft.com  

**Policies and Rules**   
https://security.microsoft.com/threatpolicy  

**Explorer**  
https://security.microsoft.com/threatexplorer  

**Legacy Retention Policies**  
https://purview.microsoft.com/datalifecyclemanagement/exchange/retentionpolicies  

**Quarantine**  
https://security.microsoft.com/quarantine  
E-mail that sends quarantine notifications: Quarantine@messaging.microsoft.com  

**User reported messages:** https://security.microsoft.com/reportsubmission  

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
**Permanently clear previous mailbox info**  
Set-User <MAILBOX> -PermanentlyClearPreviousMailboxInfo -Confirm:$false  

**Apply mandatory properties**  
Set-Mailbox <MAILBOX> -ApplyMandatoryProperties  

**Configure inbound connector to skip IPs**  
Set-InboundConnector "Relay" -EFSkipIPS 127.0.0.1,127.0.0.2  

**Show moverequest status in percentage**  
Get-Moverequest | Get-MoverequestStatistics  

**Approve migration with bad items immediately**  
Set-MoveRequest MAILBOX -SkippedItemApprovalTime $(Get-Date).ToUniversalTime()  

**Complete move request immediately**  
Set-Moverequest MAILBOX -Completeafter 1  

**Limited Details**  
Set-MailboxFolderPermission <ALIAS>:\Calendar -User Default -AccessRights LimitedDetails  
**Grant Send As permission**  
Add-RecipientPermission <MAILBOX> -Trustee <USERNAME> -AccessRights SendAs -Confirm:$false  
**Grant Full Access with Automapping**  
Add-MailboxPermission MAILBOX -User USERNAME -AccessRights FullAccess -Automapping $true  

**Retrieve message trace for the last 10 days (max 5000 results)**  
Get-MessageTraceV2 -ResultSize 5000 -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date)    

# Client Configuration (Registry)
**Hide Outlook New Button**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: HideNewOutlookToggle  
Value: 1

**Stop Outlook from becoming Outlook NEW**  
Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General  
DWORD: DoNewOutlookAutoMigration  
Value: 0  
