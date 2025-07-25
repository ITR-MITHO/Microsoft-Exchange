<#

    Optimizing security for Exchange Online mailboxes
    -	Inbound Anti-Spam Settings
     	https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-anti-spam-policy-settings
      
    -	Outbound Anti-Spam Settings
    	https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-outbound-spam-policy-settings
     
    -	Anti-Phish Settings
    	https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#eop-anti-phishing-policy-settings
     
    -	Safe Attachments
    	https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#safe-attachments-settings
    
    -	Safe Links
    	https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365#safe-links-policy-settings

    		- Mailtips
    		- Audit Log
    		- External Senders Notification
		- Safe Attachments
    		- Safe Links
    		- Anti-Phishing Policy
    		- Inbound Anti-Spam Policy
		- Outbound Anti-Spam Policy
    		- Custom Quarantine Policies, that enables notifications on RequestOnly and FullAccess
	    

#>
# Enable MailTips, Audit Log and Notify users about External Senders
Set-OrganizationConfig -MailTipsAllTipsEnabled $true -MailTipsExternalRecipientsTipsEnabled $true -MailTipsGroupMetricsEnabled $true -MailTipsLargeAudienceThreshold '25' -AuditDisabled $false | Out-null
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
Set-ExternalInOutlook -Enabled $true

# Request Only Policy
# Allows end-user to request an e-mail to be released or unblocked. The end-user receives an e-mail about quarantined e-mails.
$RequestOnly = @{
	Name					= 'ITM8 - RequestOnlyPolicy'
 	EndUserSpamNotificationFrequency	= '1.00:00:00'
  	EndUserSpamNotificationLanguage		= 'Default'
   	ESNEnabled				= $true
    	IncludeMessagesFromBlockedSenderAddress = $false
    	QuarantinePolicyType			= 'QuarantinePolicy'
	EndUserQuarantinePermissionsValue	= '43' 	
  }
New-QuarantinePolicy @RequestOnly | Out-Null

# Full Access Policy
# Allows the end-user to release e-mails or unblock senders. The end-user receives an e-mail about quarantined e-mails.
$FullAccess = @{
	Name					= 'ITM8 - FullAccessPolicy'
 	EndUserSpamNotificationFrequency	= '1.00:00:00'
  	EndUserSpamNotificationLanguage		= 'Default'
   	ESNEnabled				= $true
    	IncludeMessagesFromBlockedSenderAddress = $false
    	QuarantinePolicyType			= 'QuarantinePolicy'
	EndUserQuarantinePermissionsValue	= '39'
  }
New-QuarantinePolicy @FullAccess | Out-Null

# Admin Only Policy - With Notification
# Administrators will be notified about e-mails quarantined that the user cannot request/unblock from quarantine.
$AdminOnly = @{
	Name					= 'ITM8 - AdminOnlyPolicy'
 	EndUserSpamNotificationFrequency	= '1.00:00:00'
  	EndUserSpamNotificationLanguage		= 'Default'
   	ESNEnabled				= $false
    	IncludeMessagesFromBlockedSenderAddress = $false
    	QuarantinePolicyType			= 'QuarantinePolicy'
	EndUserQuarantinePermissionsValue	= '0'
  }
New-QuarantinePolicy @AdminOnly | Out-Null

Write-Host "
ITM8 - Quarantine policies created." -ForeGroundColor Green

Write-Host '
Mailtips, External in Outlook and Auditlogging enabled' -ForegroundColor Green

# Safe Attachment Policy for Exchange, Sharepoint and Teams
New-SafeAttachmentPolicy -Name 'ITM8 - Safe Attachments' -Action Block -Enable $true | Out-Null
New-SafeAttachmentRule -Name 'ITM8 - Safe Attachments' -SafeAttachmentPolicy 'ITM8 - Safe Attachments'  -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true | Out-Null
Set-AtpPolicyForO365 -EnableATPForSPOTeamsODB $true -EnableSafeDocs $true -AllowSafeDocsOpen $false | Out-Null

Write-Host '
ITM8 - Safe Attachments created.' -ForegroundColor Green

# New Safe Links Policy
$SafeLinks = @{
	Name = 'ITM8 - Safe Links Policy'
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
New-SafeLinksPolicy @Safelinks | Out-Null
New-SafeLinksRule -Name 'ITM8 - Safe Links Policy' -SafeLinksPolicy 'ITM8 - Safe Links Policy' -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0 -Enabled $true | Out-Null

Write-Host '
ITM8 - Safe Links Policy Created' -ForeGroundColor Green

# New Anti-Phishing Policy
$AntiPhish = @{
	Name 					= 'ITM8 - Anti-Phishing policy'
    	AdminDisplayName 			= 'ITM8 - Anti-Phishing policy'
     	EnableSpoofIntelligence			= $true
        HonorDmarcPolicy                   	= $true
    	DmarcQuarantineAction              	= 'Quarantine'
    	DmarcRejectAction                   	= 'Reject'
	AuthenticationFailAction		= 'MoveToJmf'
    	SpoofQuarantineTag                 	= 'ITM8 - RequestOnlyPolicy'
    	EnableFirstContactSafetyTips		= $true
	EnableUnauthenticatedSender 		= $true
	EnableViaTag 				= $true
    	PhishThresholdLevel             	= 3
	EnableTargetedUserProtection		= $true
	EnableOrganizationDomainsProtection	= $true
	EnableMailboxIntelligence		= $true
	EnableMailboxIntelligenceProtection	= $true
	TargetedUserProtectionAction		= 'Quarantine'
 	TargetedUserQuarantineTag		= 'ITM8 - RequestOnlyPolicy'
	TargetedDomainProtectionAction		= 'Quarantine'
	TargetedDomainQuarantineTag		= 'ITM8 - RequestOnlyPolicy'
	MailboxIntelligenceProtectionAction	= 'MoveToJmf'
	MailboxIntelligenceQuarantineTag	= 'ITM8 - RequestOnlyPolicy'
	EnableSimilarUsersSafetyTips 		= $true
	EnableSimilarDomainsSafetyTips 		= $true
	EnableUnusualCharactersSafetyTips 	= $true
}	
New-AntiPhishPolicy @AntiPhish | Out-Null
New-AntiPhishRule -Name 'ITM8 - Anti-Phishing policy' -AntiPhishPolicy 'ITM8 - Anti-Phishing policy' -RecipientDomainIs (Get-AcceptedDomain).Name -Enabled $false -Priority 0 | Out-Null

Write-Host '
ITM8 - Anti-Phishing policy created.' -ForeGroundColor Green

# New Inbound Anti-Spam Policy
$AntiSpam = @{
    	Name                                	 = 'ITM8 - Inbound Anti-Spam policy'
	IncreaseScoreWithImageLinks		 = 'Off'
	IncreaseScoreWithNumericIps		 = 'Off'
	IncreaseScoreWithRedirectToOtherPort	 = 'Off'
	IncreaseScoreWithBizOrInfoUrls		 = 'Off'
	MarkAsSpamBulkMail			 = 'On'
	MarkAsSpamEmptyMessages			 = 'Off'
	MarkAsSpamEmbedTagsInHtml 		 = 'Off'
	MarkAsSpamFormTags 			 = 'On'
	MarkAsSpamFrames 			 = 'On'
	MarkAsSpamJavaScript 			 = 'Off'
	MarkAsSpamWebBugsInHtml 		 = 'Off'
	MarkAsSpamObjectTags 			 = 'On'
	MarkAsSpamSensitiveWordList 		 = 'Off'
	MarkAsSpamSpfRecordHardFail 		 = 'Off'
	MarkAsSpamFromAddressAuthFail 		 = 'Off'
	MarkAsSpamNdrBackscatter 		 = 'Off'
	BulkThreshold 				 = '6'
 	SpamAction		                 = 'MoveToJmf'
	SpamQuarantineTag 			 = 'ITM8 - FullAccessPolicy'
 	HighConfidenceSpamAction		 = 'Quarantine'
	HighConfidenceSpamQuarantineTag 	 = 'ITM8 - FullAccessPolicy'
   	PhishSpamAction				 = 'Quarantine'
	PhishQuarantineTag 			 = 'ITM8 - FullAccessPolicy'
	HighConfidencePhishQuarantineTag 	 = 'ITM8 - AdminOnlyPolicy'
	BulkSpamAction	                    	 = 'MoveToJmf'	
	BulkQuarantineTag 			 = 'ITM8 - AdminOnlyPolicy'
	QuarantineRetentionPeriod 		 = '30'
 	EnableLanguageBlockList 		 = $false
}
New-HostedContentFilterPolicy @AntiSpam | Out-Null
New-HostedContentFilterRule -Name 'ITM8 - Inbound Anti-Spam policy' -HostedContentFilterPolicy 'ITM8 - Inbound Anti-Spam policy' -RecipientDomainIs (Get-AcceptedDomain).Name -Enabled $false | Out-Null

Write-Host '
ITM8 - Inbound Anti-Spam policy created.' -ForegroundColor Green

# New Outbound Anti-Spam Policy
$Outbound = @{
	Name					= 'ITM8 - Outbound Anti-Spam policy'
 	RecipientLimitExternalPerHour		= '500'
  	RecipientLimitInternalPerHour		= '1000'
   	RecipientLimitPerDay			= '1000'
    	ActionWhenThresholdReached		= 'BlockUser'
     	AutoForwardingMode			= 'Off'
      	BccSuspiciousOutboundMail 		= $false
       	NotifyOutboundSpam 			= $false
}
New-HostedOutboundSpamFilterPolicy @Outbound | Out-Null
New-HostedOutboundSpamFilterRule -Name 'ITM8 - Outbound Anti-Spam policy' -HostedOutboundSpamFilterPolicy 'ITM8 - Outbound Anti-Spam policy' -SenderDomainIs (Get-AcceptedDomain).Name -Enabled $false | Out-Null

Write-Host '
ITM8 - Outbound Anti-Spam policy created.' -ForegroundColor Green

Write-Host '

IMPORTANT: Anti-spam inbound/outbound and Phishing Policies are all DISABLED. They need to be enabled before they take effect' -ForeGroundColor Yellow
