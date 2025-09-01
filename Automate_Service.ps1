<#

ITM8 Exchange Server - STANDARD
This script sets our baseline/standard settings for a Exchange Online environment

MailTips
Safe Attachment Policy
Safe Links Policy
LimitedDetails as default calendar permission


#>

$Customer = Import-csv $home\Desktop\Customers.csv
Foreach ($C in $Customer)
{

# Connect to Exchange Online
Connect-ExchangeOnline -AppID $C.App -CertificateThumbprint $C.Thumb -Organization $C.Org

# MailTips
Set-OrganizationConfig -MailTipsAllTipsEnabled $true -MailTipsExternalRecipientsTipsEnabled $true -MailTipsGroupMetricsEnabled $true -MailTipsLargeAudienceThreshold '25' -AuditDisabled $false | Out-null
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
Set-ExternalInOutlook -Enabled $true

# Safe Attachment Policy for Exchange, Sharepoint and Teams
New-SafeAttachmentPolicy -Name 'ITM8 - Safe Attachments' -Action Block -Enable $true | Out-Null
New-SafeAttachmentRule -Name 'ITM8 - Safe Attachments' -SafeAttachmentPolicy 'ITM8 - Safe Attachments'  -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true | Out-Null
Set-AtpPolicyForO365 -EnableATPForSPOTeamsODB $true -EnableSafeDocs $true -AllowSafeDocsOpen $false | Out-Null

# New Safe Links Policy
$SafeLinks = @{
	Name = 'ITM8 - Safe Links Policy'
	EnableSafeLinksForEmail		    = $true
	EnableForInternalSenders	    = $true
	ScanUrls			            = $true
	DeliverMessageAfterScan		    = $true
	DisableUrlRewrite		        = $false
	EnableSafeLinksForTeams 	    = $true
	EnableSafeLinksForOffice 	    = $true
	TrackClicks 			        = $true
	AllowClickThrough		        = $false
    EnableOrganizationBranding  	= $false
    UseTranslatedNotificationText   = $false
}
New-SafeLinksPolicy @Safelinks | Out-Null
New-SafeLinksRule -Name 'ITM8 - Safe Links Policy' -SafeLinksPolicy 'ITM8 - Safe Links Policy' -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true | Out-Null


# Default UserMailbox Calendar Permissions
$User = 'Default'
$AccessRight = 'LimitedDetails'
Foreach ($Mailbox in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{
    $UserPrincipalName = $Mailbox.UserPrincipalName
    $Calendar = (Get-MailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -FolderScope Calendar | Where { $_.FolderType -eq 'Calendar'}).Name
Try 
{
    
    Set-MailboxFolderPermission -Identity ($Mailbox.UserPrincipalName+":\$Calendar") -User $User -AccessRights $AccessRight -WarningAction SilentlyContinue -ErrorAction Stop

}
Catch
{
    Write-Warning "Failed to add the user '$User' with calendar permission '$AccessRight' on Mailbox: $UserPrincipalName"
    Continue
}
    }
        }
