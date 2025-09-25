<#

ITM8 Exchange Service - STANDARD
This script sets our baseline/standard settings for a Exchange Online environment

ExternalInOutlook - Is set for everyone
LimitedDetails - Is set as the default calendar permission on UserMailboxes
RoomMailbox - A set of default parameters
SharedMailbox - Save sent items in the mailbox
User and Shared - Max Send/Receive 150MB and RetainDeletedItems set to 30 days
Backup - Every 15 day a export of all mailboxes attributes is exported

#>
# -DO NOT CHANGE THESE VALUES-
$ResourceGroup = "rg-exchange"
$AutomationAccount = 'AA-Exchange'

# Change these to fit the customer specific information.
$CertThumb = 'CERTIFICATETHUMBPRINT'
$AppID = 'APPID'
$OrganizationName = 'domain-com.onmicrosoft.com'

# Change these variables to match the customers wishes
$RetainDeletedItems = '30.00:00:00'
$Receive = '150MB'
$Send = '150MB'
$ExternalInOutlook = $true
$RetentionPolicy = 'ITM8 - Deleted Items - 30 days'
$CalPer = 'Reviewer' # This value is to set the permissions for 'default' on all calendars

$Start = Get-Date
# Import Exchange Online Module
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Exchange Online using the Certificate Thumbprint of the Certificate imported into the Automation Account
Try {
Connect-ExchangeOnline -CertificateThumbPrint $CertThumb -AppID "$AppID" -Organization "$OrganizationName" -ErrorAction Stop
Write-Output "$Start - Connected to Exchange Online"
    }
Catch
{
    Write-Output 'Failed to connect to Exchange Online for $OrganizationName'
}

# Start of the actual script; 
$Mailbox = Get-Mailbox -Resultsize Unlimited
$Count = ($Mailbox).Count

# Nice to have settings
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true | Out-Null
Set-ExternalInOutlook -Enabled $ExternalInOutlook | Out-Null

# Sent items in original folder
$Mailbox | Where {$_.RecipientTypeDetails -EQ 'SharedMailbox'} | Set-Mailbox -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true

# Max Send/Receive Size
Foreach ($M in $Mailbox)
{
Set-Mailbox -Identity $M.ExchangeGuid -MaxSendSize $Send -MaxReceiveSize $Receive -RetentionPolicy $RetentionPolicy -RetainDeletedItemsFor $RetainDeletedItems
}

# Calendar Permissions
Foreach ($M in $Mailbox)
{
    $Calendar = (Get-MailboxFolderStatistics -Identity $M.ExchangeGuid -FolderScope Calendar | Where { $_.FolderType -eq 'Calendar'}).Name
Try 
{
    Set-MailboxFolderPermission -Identity ($M.ExchangeGuid+":\$Calendar") -User Default -AccessRights $CalPer -WarningAction SilentlyContinue -ErrorAction Stop
}
Catch
{
    Write-Warning "Failed to add the calendar permission '$AccessRight' on Mailbox: $UserPrincipalName"
    Continue
}
    }
$End = Get-Date
Write-Output "$End - Changes made to $Count Mailboxes"
