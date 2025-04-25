<#

    Optimizing security for Exchange Online mailboxes
    Anti-Spam Settings: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-anti-spam-policy-settings
    Anti-Phish Settings: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-anti-phishing-policy-settings
    Safe Attachments: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#safe-attachments-settings
    Safe Links: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#safe-links-policy-settings
    
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
Set-AtpPolicyForO365 -EnableATPForSPOTeamsODB $true -EnableSafeDocs $true -AllowSafeDocsOpen $false

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
    	EnableOrganizationBranding  	= $false
	DeliverMessageAfterScan		= $true
	DisableUrlRewrite		= $false
}
New-SafeLinksPolicy @Safelinks
New-SafeLinksRule -Name "ITM8 - Safe Links Policy" -SafeLinksPolicy "ITM8 - Safe Links Policy" -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true

# New Anti-Phishing Policy
$AntiPhish = @{
	Name 					= "ITM8 - Anti-Phishing policy"
    	AdminDisplayName 			= "ITM8 - Anti-Phishing policy"
    	PhishThresholdLevel             	= 3
	EnableTargetedUserProtection		= $true
	EnableOrganizationDomainsProtection	= $true
	EnableMailboxIntelligence		= $true
	EnableMailboxIntelligenceProtection	= $true
	EnableSpoofIntelligence			= $true
	TargetedUserProtectionAction		= "Quarantine"
	TargetedDomainProtectionAction		= "Quarantine"
	MailboxIntelligenceProtectionAction	= "MoveToJmf"
	AuthenticationFailAction		= "MoveToJmf"
    	DmarcQuarantineAction              	= "Quarantine"
    	DmarcRejectAction                   	= "Reject"
    	EnableFirstContactSafetyTips		= $true
	EnableSimilarUsersSafetyTips 		= $true
	EnableSimilarDomainsSafetyTips 		= $true
	EnableUnusualCharactersSafetyTips 	= $true
	EnableUnauthenticatedSender 		= $true
	EnableViaTag 				= $true
    	HonorDmarcPolicy                   	= $true
    	SpoofQuarantineTag                 	= "DefaultFullAccessPolicy"
	TargetedDomainQuarantineTag		= "DefaultFullAccessWithNotificationPolicy"
	MailboxIntelligenceQuarantineTag	= "DefaultFullAccessPolicy"
	TargetedUserQuarantineTag		= "DefaultFullAccessWithNotificationPolicy"
}	
New-AntiPhishPolicy @AntiPhish
New-AntiPhishRule -Name "ITM8 - Anti-Phishing policy" -AntiPhishPolicy "ITM8 - Anti-Phishing policy" -RecipientDomainIs (Get-AcceptedDomain).Name -Enabled $true -Priority 0

# New Inbound Anti-Spam Policy
$AntiSpam = @{
    	Name                                	 = "ITM8 - Inbound Anti-Spam policy"
	IncreaseScoreWithImageLinks		 = "Off"
	IncreaseScoreWithNumericIps		 = "Off"
	IncreaseScoreWithRedirectToOtherPort	 = "Off"
	IncreaseScoreWithBizOrInfoUrls		 = "Off"
	MarkAsSpamBulkMail			 = "On"
	MarkAsSpamEmptyMessages			 = "Off"
	MarkAsSpamEmbedTagsInHtml 		 = "Off"
	MarkAsSpamFormTags 			 = "On"
	MarkAsSpamFrames 			 = "On"
	MarkAsSpamJavaScript 			 = "Off"
	MarkAsSpamWebBugsInHtml 		 = "Off"
	MarkAsSpamObjectTags 			 = "On"
	MarkAsSpamSensitiveWordList 		 = "Off"
	MarkAsSpamSpfRecordHardFail 		 = "Off"
	MarkAsSpamFromAddressAuthFail 		 = "Off"
	MarkAsSpamNdrBackscatter 		 = "Off"
	BulkThreshold 				 = "6"
 	SpamAction		                 = "MoveToJmf"
	SpamQuarantineTag 			 = "DefaultFullAccessPolicy"
 	HighConfidenceSpamAction		 = "Quarantine"
	HighConfidenceSpamQuarantineTag 	 = "DefaultFullAccessWithNotificationPolicy"
   	PhishSpamAction				 = "Quarantine"
	PhishQuarantineTag 			 = "DefaultFullAccessWithNotificationPolicy"
	HighConfidencePhishQuarantineTag 	 = "AdminOnlyAccessPolicy"
	BulkSpamAction	                    	 = "MoveToJmf"	
	BulkQuarantineTag 			 = "DefaultFullAccessPolicy"
	QuarantineRetentionPeriod 		 = "30"
 	EnableLanguageBlockList 		 = $false
}
New-HostedContentFilterPolicy @AntiSpam
New-HostedContentFilterRule -Name "ITM8 - Inbound Anti-Spam policy" -HostedContentFilterPolicy "ITM8 - Inbound Anti-Spam policy" -RecipientDomainIs (Get-AcceptedDomain).Name

