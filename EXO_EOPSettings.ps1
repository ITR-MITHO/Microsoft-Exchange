<#

    Optimizing security for Exchange Online mailboxes
    Based on: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365
    
    - Mailtips
    - Audit Log
    - External Senders Notification
    - Safe Attachments
    - Safe Links
    - Anti-Phishing Policy
    - Anti-Spam Policy

#>
# Enable MailTips, Audit Log and Notify users about External Senders
Set-OrganizationConfig -MailTipsAllTipsEnabled $true -MailTipsExternalRecipientsTipsEnabled $true -MailTipsGroupMetricsEnabled $true -MailTipsLargeAudienceThreshold '25' -AuditDisabled $false
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
Set-ExternalInOutlook –Enabled $true

# Safe Attachment Policy for Exchange, Sharepoint and Teams
New-SafeAttachmentPolicy -Name "ITM8 - Safe Attachments" -Action Block -Enable $true
New-SafeAttachmentRule -Name "ITM8 - Safe Attachments" -SafeAttachmentPolicy "ITM8 - Safe Attachments"  -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true
Set-AtpPolicyForO365 -EnableATPForSPOTeamsODB $true

# New Safe Links Policy
$SafeLinks = @{
	Name = "ITM8 - Safe Links Policy"
	EnableSafeLinksForEmail		= $true
	EnableSafeLinksForTeams 	= $true
	EnableSafeLinksForOffice 	= $true
	TrackClicks 			= $true
	AllowClickThrough		= $false
	ScanUrls			= $true
	EnableForInternalSenders	= $true
    	EnableOrganizationBranding  	= $true
	DeliverMessageAfterScan		= $true
	DisableUrlRewrite		= $false
}
New-SafeLinksPolicy @Safelinks
# Create the rule for all users in all valid domains and associate with policy.
New-SafeLinksRule -Name "ITM8 - Safe Links Policy" -SafeLinksPolicy "ITM8 - Safe Links Policy" -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true

# New anti-phishing policy
$AntiPhish = @{
	Name = "ITM8 - Anti-Phishing policy"
    	AdminDisplayName 			= "ITM8 - Anti-Phishing policy"
    	PhishThresholdLevel             	= 3
	EnableTargetedUserProtection		= $true
	EnableOrganizationDomainsProtection	= $true
	EnableMailboxIntelligence		= $true
	EnableMailboxIntelligenceProtection	= $true
	EnableSpoofIntelligence			= $true
}
New-AntiPhishPolicy @AntiPhish

# Anti-Phish Settings
$Actions = @{
	TargetedUserProtectionAction		= "Quarantine"
	TargetedDomainProtectionAction		= "Quarantine"
	MailboxIntelligenceProtectionAction	= "Quarantine"
	AuthenticationFailAction		= "MoveToJmf"
    	DmarcQuarantineAction              	= "Quarantine"
    	DmarcRejectAction                   	= "Reject"
    	EnableFirstContactSafetyTips		= $true
	EnableSimilarUsersSafetyTips 		= $true
	EnableSimilarDomainsSafetyTips 		= $true
	EnableUnusualCharactersSafetyTips 	= $true
	EnableUnauthenticatedSender 		= $true
	EnableViaTag 				= $true
    	EnableSpoofIntelligence             	= $true
    	HonorDmarcPolicy                    	= $true
    	SpoofQuarantineTag                  	= "DefaultFullAccessPolicy"

}
Set-AntiPhishPolicy -Identity "ITM8 - Anti-Phishing policy" @Actions
New-AntiPhishRule -Name "ITM8 - Anti-Phishing policy" -AntiPhishPolicy "ITM8 - Anti-Phishing policy" -RecipientDomainIs (Get-AcceptedDomain).Name -Enabled $true -Priority 0


# New Anti-spam inbound policy
New-HostedContentFilterPolicy -Name "ITM8 - Inbound Anti-Spam policy" `
    -SpamAction Quarantine `
    -HighConfidenceSpamAction Quarantine `
    -BulkSpamAction Quarantine `
    -EnableEndUserSpamNotifications $true `
    -EndUserSpamNotificationFrequency 3

# Set Anti-Spam Settings
Set-HostedContentFilterPolicy -Identity "ITM8 - Inbound Anti-Spam policy" `
    -IncreaseScoreWithImageLinks "Off" `
    -IncreaseScoreWithNumericIps "Off" `
    -IncreaseScoreWithRedirectToOtherPort "Off" `
    -IncreaseScoreWithBizOrInfoUrls "Off" `
    -MarkAsSpamBulkMail "On" `
    -MarkAsSpamEmptyMessages "Off" `
    -MarkAsSpamEmbedTagsInHtml "Off" `
    -MarkAsSpamFormTags "On" `
    -MarkAsSpamFrames "On" `
    -MarkAsSpamJavaScript "Off" `
    -MarkAsSpamWebBugsInHtml "Off" `
    -MarkAsSpamObjectTags "On" `
    -MarkAsSpamSensitiveWordList "Off" `
    -MarkAsSpamSpfRecordHardFail "Off" `
    -MarkAsSpamFromAddressAuthFail "Off" `
    -MarkAsSpamNdrBackscatter "Off" `
    -BulkThreshold 6 `
    -SpamQuarantineTag DefaultFullAccessPolicy `
    -HighConfidenceSpamQuarantineTag DefaultFullAccessPolicy `
    -BulkQuarantineTag DefaultFullAccessPolicy `
    -PhishQuarantineTag AdminOnlyAccessPolicy `
    -HighConfidencePhishQuarantineTag AdminOnlyAccessPolicy `
    -QuarantineRetentionPeriod 30

# Apply the policy to domains
New-HostedContentFilterRule -Name "ITM8 - Inbound Anti-Spam policy" -HostedContentFilterPolicy "ITM8 - Inbound Anti-Spam policy" -RecipientDomainIs (Get-AcceptedDomain).Name
