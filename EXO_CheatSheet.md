# Spam settings and quarantine:
https://security.microsoft.com/threatpolicy

https://security.microsoft.com/quarantine

# Legacy Retention Policies:
https://purview.microsoft.com/datalifecyclemanagement/exchange/retentionpolicies

# Move requests
Get-Moverequest | Get-MoverequestStatistics

Set-MoveRequest Mailbox1 -SkippedItemApprovalTime $(Get-Date).ToUniversalTime()

Set-Moverequest Mailbox1 -Completeafter 1

# Permissions
Set-MailboxFolderPermission USER:\Calendar -User Default -AccessRights LimitedDetails

Add-RecipientPermission -Identity SharedMailbox -Trustee User -AccessRights Sendas -Confirm:$false

Add-MailboxPermission -Identity Mailbox -User Username -AccessRights FullAccess -Automapping $true

# Nice to know
Set-User Jon@contoso.com -PermanentlyClearPreviousMailboxInfo -confirm:$false

Set-Mailbox Mailbox1 -ApplyMandatoryProperties

Get-MessageTrace -Start (Get-date).AddDays(-10) -End (Get-Date)

# Mailflow
SPF: spf.protection.outlook.com

MX: domain-dk.mail.protection.outlook.com

MX: domain-dk.l-v1.mx.microsoft

# Network
Port: 25

Source/Destination:
*.mail.protection.outlook.com, *.mx.microsoft
40.92.0.0/15, 40.107.0.0/16, 52.100.0.0/14, 104.47.0.0/17, 2a01:111:f400::/48, 2a01:111:f403::/48

Port: 443

Source/Destination:
*.protection.outlook.com
40.92.0.0/15, 40.107.0.0/16, 52.100.0.0/14, 52.238.78.88/32, 104.47.0.0/17, 2a01:111:f400::/48, 2a01:111:f403::/48

# Limits
https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits

# Outlook NEW
**Hide Outlook New Button**

Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General

DWORD: HideNewOutlookToggle

Value: 00000000


**Stop automatic Outlook New Upgrade**

Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General

DWORD: DoNewOutlookAutoMigration

Value: 00000000
