<#

    Optimizing security for Exchange Online mailboxes
    Inbound Anti-Spam Settings: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-anti-spam-policy-settings
    Outbound Anti-Spam Settings: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-outbound-spam-policy-settings
    Anti-Phish Settings: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-anti-phishing-policy-settings
    Safe Attachments: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#safe-attachments-settings
    Safe Links: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#safe-links-policy-settings
    
    - Mailtips
    - Audit Log
    - External Senders Notification
    - Safe Attachments
    - Safe Links
    - Anti-Phishing Policy
    - Inbound Anti-Spam Policy
    - Outbound Anti-Spam Policy
    - Perhaps... Custom Quarantine Policies, that enables notifications on all 3 policies?

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
	EnableForInternalSenders	= $true
	ScanUrls			= $true
	DeliverMessageAfterScan		= $true
	DisableUrlRewrite		= $false
	EnableSafeLinksForTeams 	= $true
	EnableSafeLinksForOffice 	= $true
	TrackClicks 			= $true
	AllowClickThrough		= $false
    	EnableOrganizationBranding  	= $false
        UseTranslatedNotificationText   = $false
}
New-SafeLinksPolicy @Safelinks
New-SafeLinksRule -Name "ITM8 - Safe Links Policy" -SafeLinksPolicy "ITM8 - Safe Links Policy" -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true

# New Anti-Phishing Policy
$AntiPhish = @{
	Name 					= "ITM8 - Anti-Phishing policy"
    	AdminDisplayName 			= "ITM8 - Anti-Phishing policy"
     	EnableSpoofIntelligence			= $true
        HonorDmarcPolicy                   	= $true
    	DmarcQuarantineAction              	= "Quarantine"
    	DmarcRejectAction                   	= "Reject"
	AuthenticationFailAction		= "MoveToJmf"
    	SpoofQuarantineTag                 	= "DefaultFullAccessPolicy"
    	EnableFirstContactSafetyTips		= $true
	EnableUnauthenticatedSender 		= $true
	EnableViaTag 				= $true
    	PhishThresholdLevel             	= 3
	EnableTargetedUserProtection		= $true
	EnableOrganizationDomainsProtection	= $true
	EnableMailboxIntelligence		= $true
	EnableMailboxIntelligenceProtection	= $true
	TargetedUserProtectionAction		= "Quarantine"
 	TargetedUserQuarantineTag		= "DefaultFullAccessWithNotificationPolicy"
	TargetedDomainProtectionAction		= "Quarantine"
	TargetedDomainQuarantineTag		= "DefaultFullAccessWithNotificationPolicy"
	MailboxIntelligenceProtectionAction	= "MoveToJmf"
	MailboxIntelligenceQuarantineTag	= "DefaultFullAccessPolicy"
	EnableSimilarUsersSafetyTips 		= $true
	EnableSimilarDomainsSafetyTips 		= $true
	EnableUnusualCharactersSafetyTips 	= $true
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

# New Outbound Anti-Spam Policy
$Outbound = @{
	Name					= "ITM8 - Outbound Anti-Spam policy"
 	RecipientLimitExternalPerHour		= "500"
  	RecipientLimitInternalPerHour		= "100"
   	RecipientLimitPerDay			= "1000"
    	ActionWhenThresholdReached		= "BlockUser"
     	AutoForwardingMode			= "Off"
      	BccSuspiciousOutboundMail 		= $false
       	NotifyOutboundSpam 			= $false
}
New-HostedOutboundSpamFilterPolicy @Outbound
New-HostedOutboundSpamFilterRule -Name "ITM8 - Outbound Anti-Spam policy" -HostedOutboundSpamFilterPolicy "ITM8 - Outbound Anti-Spam policy" -SenderDomainIs (Get-AcceptedDomain).Name

<# Maybe... Maybe not? Quarantine policies, where all have notifications enabled??

Read Only Policy - With Notification
$ReadOnly = @{
	Name					= "ITM8 - ReadOnlyPolicy"
 	EndUserSpamNotificationFrequency	= "1.00:00:00"
  	EndUserSpamNotificationLanguage		= "Default"
   	ESNEnabled				= $true
    	IncludeMessagesFromBlockedSenderAddress = $false
    	QuarantinePolicyType			= "QuarantinePolicy"
	EndUserQuarantinePermissionsValue	= "43" 	
  }
New-QuarantinePolicy @ReadOnly

# Full Access Policy - With Notification
$FullAccess = @{
	Name					= "ITM8 - FullAccessPolicy"
 	EndUserSpamNotificationFrequency	= "1.00:00:00"
  	EndUserSpamNotificationLanguage		= "Default"
   	ESNEnabled				= $true
    	IncludeMessagesFromBlockedSenderAddress = $false
    	QuarantinePolicyType			= "QuarantinePolicy"
	EndUserQuarantinePermissionsValue	= "39"
  }
New-QuarantinePolicy @FullAccess

# Admin Only Policy - With Notification
$AdminOnly = @{
	Name					= "ITM8 - AdminOnlyPolicy"
 	EndUserSpamNotificationFrequency	= "1.00:00:00"
  	EndUserSpamNotificationLanguage		= "Default"
   	ESNEnabled				= $true
    	IncludeMessagesFromBlockedSenderAddress = $false
    	QuarantinePolicyType			= "QuarantinePolicy"
	EndUserQuarantinePermissionsValue	= "0"
  }
New-QuarantinePolicy @AdminOnly
<#
