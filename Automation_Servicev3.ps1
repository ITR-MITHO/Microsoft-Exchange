# DO NOT CHANGE THESE VALUES
$ResourceGroup = "rg-exchange-configservice-001"
$AutomationAccount = 'AA-Exchange'

# Change these to fit the customer specific information.
$Certname = 'MYCERT'
$AppID = 'APPID'
$OrganizationName = 'XXXXX.onmicrosoft.com'
$Start = Get-Date

# Import Exchange Online Module
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Exchange Online using the Certificate Thumbprint of the Certificate imported into the Automation Account
Try {
Connect-ExchangeOnline -CertificateThumbPrint 'XXXXXXXXXXXXXXXXXXXXXXXXXX' -AppID "$AppID" -Organization "$OrganizationName" -ErrorAction Stop
Write-Output "$Start - Connected to Exchange Online"
    }
Catch
{
    Write-Output 'Failed to connect to Exchange Online for $OrganizationName'
}

# Start of the actual script; 
$Mailbox = Get-Mailbox -ResultSize Unlimited
$Count = ($Mailbox).Count

# Nice to have settings
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true | Out-Null
Set-ExternalInOutlook -Enabled $true | Out-Null

# Send items in original folder
$Mailbox | Where {$_.RecipientTypeDetails -EQ 'SharedMailbox'} | Set-Mailbox -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true

# Max Send/Receive Size
Foreach ($M in $Mailbox)
{
Set-Mailbox -Identity $M.Alias -MaxSendSize '150MB' -MaxReceiveSize '150MB' -RetentionPolicy 'ITM8 - Deleted Items - 30 days'
}

# Calendar Permissions
$User = 'Default'
$AccessRight = 'Reviewer'
Foreach ($M in $Mailbox)
{
    $UserPrincipalName = $Cal.UserPrincipalName
    $Calendar = (Get-MailboxFolderStatistics -Identity $M.UserPrincipalName -FolderScope Calendar | Where { $_.FolderType -eq 'Calendar'}).Name
Try 
{
    Set-MailboxFolderPermission -Identity ($M.UserPrincipalName+":\$Calendar") -User $User -AccessRights $AccessRight -WarningAction SilentlyContinue -ErrorAction Stop
}
Catch
{
    Write-Warning "Failed to add the user '$User' with calendar permission '$AccessRight' on Mailbox: $UserPrincipalName"
    Continue
}
    }
$End = Get-Date
Write-Output "$End - Changes made to $Count Mailboxes"
