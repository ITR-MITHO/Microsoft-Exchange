<#
   Script to setup Microsoft EOP Recommendation baselines. 
   The script follows the settings mentioned by Microsoft in this article: https://learn.microsoft.com/en-us/defender-office-365/recommended-settings-for-eop-and-office365
#>

$ErrorActionPreference = "Stop"
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
} catch {
    Write-Warning "Not connected to Exchange Online. Please run Connect-ExchangeOnline."
    return
}
$AcceptedDomains = (Get-AcceptedDomain).Name

# Global Configurations
Write-Host "Configuring Global Settings (MailTips, Audit Log, External Senders)..." -ForegroundColor Cyan
try {
    $null = Set-OrganizationConfig -MailTipsAllTipsEnabled $true -MailTipsExternalRecipientsTipsEnabled $true -MailTipsGroupMetricsEnabled $true -MailTipsLargeAudienceThreshold 25 -AuditDisabled $false
    $null = Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -ErrorAction SilentlyContinue # May throw warning in newer tenants as it is default
    $null = Set-ExternalInOutlook -Enabled $true
    Write-Host "Global settings configured successfully." -ForegroundColor Green
} catch {
    Write-Error "Failed to configure global settings: $_"
}

# Quarantine Policies
$QuarantinePolicies = @(
    @{
        Name = 'ITM8 - RequestOnlyPolicy'
        EndUserSpamNotificationFrequency = '1.00:00:00'
        EndUserSpamNotificationLanguage = 'Default'
        ESNEnabled = $true
        IncludeMessagesFromBlockedSenderAddress = $false
        QuarantinePolicyType = 'QuarantinePolicy'
        EndUserQuarantinePermissionsValue = 43
    },
    @{
        Name = 'ITM8 - FullAccessPolicy'
        EndUserSpamNotificationFrequency = '1.00:00:00'
        EndUserSpamNotificationLanguage = 'Default'
        ESNEnabled = $true
        IncludeMessagesFromBlockedSenderAddress = $false
        QuarantinePolicyType = 'QuarantinePolicy'
        EndUserQuarantinePermissionsValue = 7
    },
    @{
        Name = 'ITM8 - AdminOnlyPolicy'
        EndUserSpamNotificationFrequency = '1.00:00:00'
        EndUserSpamNotificationLanguage = 'Default'
        ESNEnabled = $false
        IncludeMessagesFromBlockedSenderAddress = $false
        QuarantinePolicyType = 'QuarantinePolicy'
        EndUserQuarantinePermissionsValue = 0
    }
)

foreach ($policy in $QuarantinePolicies) {
    if (-not (Get-QuarantinePolicy -Identity $policy.Name -ErrorAction SilentlyContinue)) {$null = New-QuarantinePolicy @policy
        Write-Host "Created Quarantine Policy: $($policy.Name)" -ForegroundColor Green
    } else {
        Write-Host "Quarantine Policy already exists: $($policy.Name)" -ForegroundColor Yellow
    }
}

# Safe Attachments (Requires Microsoft Defender for Office 365)
if (Get-Command Get-SafeAttachmentPolicy -ErrorAction SilentlyContinue) {
    $SafeAttachmentName = 'ITM8 - Safe Attachments'
    if (-not (Get-SafeAttachmentPolicy -Identity $SafeAttachmentName -ErrorAction SilentlyContinue)) {
        $null = New-SafeAttachmentPolicy -Name$SafeAttachmentName -Action Block -Enable $true 
		$null = New-SafeAttachmentRule -Name $SafeAttachmentName -SafeAttachmentPolicy $SafeAttachmentName -RecipientDomainIs $AcceptedDomains -Priority 0 -Enabled $false
        Write-Host "Created Safe Attachments Policy: $SafeAttachmentName" -ForegroundColor Green
    } else {
        Write-Host "Safe Attachments Policy already exists: $SafeAttachmentName" -ForegroundColor Yellow
    }

    if (Get-Command Set-AtpPolicyForO365 -ErrorAction SilentlyContinue) {
        $null = Set-AtpPolicyForO365 -EnableATPForSPOTeamsODB$true -EnableSafeDocs $true -AllowSafeDocsOpen$false
    }
} else {
    Write-Warning "Tenant is not licensed for Safe Attachments (Microsoft Defender for Office 365 Plan 1/2 required). Skipping."
}

# Safe Links (Requires Microsoft Defender for Office 365)
if (Get-Command Get-SafeLinksPolicy -ErrorAction SilentlyContinue) {
    $SafeLinksName = 'ITM8 - Safe Links Policy'
	$SafeLinksParams = @{
        Name = $SafeLinksName
        EnableSafeLinksForEmail = $true
        EnableForInternalSenders = $true
        ScanUrls = $true
        DeliverMessageAfterScan = $true
        DisableUrlRewrite = $false
        EnableSafeLinksForTeams = $true
        EnableSafeLinksForOffice = $true
        TrackClicks = $true
        AllowClickThrough = $false
        EnableOrganizationBranding = $false
        UseTranslatedNotificationText = $false
    }

    if (-not (Get-SafeLinksPolicy -Identity $SafeLinksName -ErrorAction SilentlyContinue)) {$null = New-SafeLinksPolicy @SafeLinksParams
        $null = New-SafeLinksRule -Name $SafeLinksName -SafeLinksPolicy$SafeLinksName -RecipientDomainIs $AcceptedDomains -Priority 0 -Enabled $false
        Write-Host "Created Safe Links Policy: $SafeLinksName" -ForegroundColor Green
    } else {
        Write-Host "Safe Links Policy already exists: $SafeLinksName" -ForegroundColor Yellow
    }
} else {
    Write-Warning "Tenant is not licensed for Safe Links (Microsoft Defender for Office 365 Plan 1/2 required). Skipping."
}

# Anti-Phishing Policy
if (Get-Command New-AntiPhishPolicy -ErrorAction SilentlyContinue) {
    $AntiPhishName = 'ITM8 - Anti-Phishing policy'

    # Base EOP parameters supported across ALL tenants
    $BaseParams = @{
        Name                               = $AntiPhishName
        AdminDisplayName                   = $AntiPhishName
        EnableSpoofIntelligence            = $true
        HonorDmarcPolicy                   = $true
        DmarcQuarantineAction              = 'Quarantine'
        DmarcRejectAction                  = 'Reject'
        AuthenticationFailAction           = 'MoveToJmf'
        SpoofQuarantineTag                 = 'ITM8 - RequestOnlyPolicy'
        EnableFirstContactSafetyTips       = $true
        EnableUnauthenticatedSender        = $true
        EnableViaTag                       = $true
    }

    # Defender for Office 365 (Plan 1/2) specific parameters
    $DefenderParams = @{
        PhishThresholdLevel                = 3
        EnableTargetedUserProtection       = $true
        EnableOrganizationDomainsProtection= $true
        EnableMailboxIntelligence          = $true
        EnableMailboxIntelligenceProtection= $true
        TargetedUserProtectionAction       = 'Quarantine'
        TargetedUserQuarantineTag          = 'ITM8 - RequestOnlyPolicy'
        TargetedDomainProtectionAction     = 'Quarantine'
        TargetedDomainQuarantineTag        = 'ITM8 - RequestOnlyPolicy'
        MailboxIntelligenceProtectionAction= 'MoveToJmf'
        MailboxIntelligenceQuarantineTag   = 'ITM8 - RequestOnlyPolicy'
        EnableSimilarUsersSafetyTips        = $true
        EnableSimilarDomainsSafetyTips       = $true
        EnableUnusualCharactersSafetyTips   = $true
    }

    $CommandMetaData = Get-Command New-AntiPhishPolicy
    $SupportsDefender = $CommandMetaData.Parameters.ContainsKey('MailboxIntelligenceQuarantineTag')

    if ($SupportsDefender) {
        foreach ($key in $DefenderParams.Keys) {
            $BaseParams[$key] = $DefenderParams[$key]
        }
    } else {
        Write-Warning "Tenant lacks Defender for O365 licensing. Applying basic EOP Anti-Phishing settings only."
    }

    if (-not (Get-AntiPhishPolicy -Identity $AntiPhishName -ErrorAction SilentlyContinue)) {
        $null = New-AntiPhishPolicy @BaseParams
        $null = New-AntiPhishRule -Name $AntiPhishName -AntiPhishPolicy $AntiPhishName -RecipientDomainIs $AcceptedDomains -Enabled $false -Priority 0
        Write-Host "Created Anti-Phishing Policy: $AntiPhishName" -ForegroundColor Green
    } else {
        Write-Host "Anti-Phishing Policy already exists: $AntiPhishName" -ForegroundColor Yellow
    }
}

# Inbound Anti-Spam Policy
$InboundSpamName = 'ITM8 - Inbound Anti-Spam policy'
$AntiSpamParams = @{
    Name = $InboundSpamName
    IncreaseScoreWithImageLinks = 'Off'
    IncreaseScoreWithNumericIps = 'Off'
    IncreaseScoreWithRedirectToOtherPort = 'Off'
    IncreaseScoreWithBizOrInfoUrls = 'Off'
    MarkAsSpamBulkMail = 'On'
    MarkAsSpamEmptyMessages = 'Off'
    MarkAsSpamEmbedTagsInHtml = 'Off'
    MarkAsSpamFormTags = 'On'
    MarkAsSpamFrames = 'On'
    MarkAsSpamJavaScript = 'Off'
    MarkAsSpamWebBugsInHtml = 'Off'
    MarkAsSpamObjectTags = 'On'
    MarkAsSpamSensitiveWordList = 'Off'
    MarkAsSpamSpfRecordHardFail = 'Off'
    MarkAsSpamFromAddressAuthFail = 'Off'
    MarkAsSpamNdrBackscatter = 'Off'
    BulkThreshold = 6
    SpamAction = 'MoveToJmf'
    SpamQuarantineTag = 'ITM8 - FullAccessPolicy'
    HighConfidenceSpamAction = 'Quarantine'
    HighConfidenceSpamQuarantineTag = 'ITM8 - FullAccessPolicy'
    PhishSpamAction = 'Quarantine'
    PhishQuarantineTag = 'ITM8 - FullAccessPolicy'
    HighConfidencePhishQuarantineTag = 'ITM8 - AdminOnlyPolicy'
    BulkSpamAction = 'MoveToJmf'
    BulkQuarantineTag = 'ITM8 - AdminOnlyPolicy'
    QuarantineRetentionPeriod = 30
    EnableLanguageBlockList = $false
}

if (-not (Get-HostedContentFilterPolicy -Identity $InboundSpamName -ErrorAction SilentlyContinue)) {$null = New-HostedContentFilterPolicy @AntiSpamParams
    $null = New-HostedContentFilterRule -Name $InboundSpamName -HostedContentFilterPolicy$InboundSpamName -RecipientDomainIs $AcceptedDomains -Enabled$false
    Write-Host "Created Inbound Anti-Spam Policy: $InboundSpamName" -ForegroundColor Green
} else {
    Write-Host "Inbound Anti-Spam Policy already exists: $InboundSpamName" -ForegroundColor Yellow
}

# Outbound Anti-Spam Policy
$OutboundSpamName = 'ITM8 - Outbound Anti-Spam policy'
$OutboundParams = @{
    Name = $OutboundSpamName
    RecipientLimitExternalPerHour = 500
    RecipientLimitInternalPerHour = 1000
    RecipientLimitPerDay = 1000
    ActionWhenThresholdReached = 'BlockUser'
    AutoForwardingMode = 'Off'
    BccSuspiciousOutboundMail = $false
    NotifyOutboundSpam = $false
}

if (-not (Get-HostedOutboundSpamFilterPolicy -Identity $OutboundSpamName -ErrorAction SilentlyContinue)) {$null = New-HostedOutboundSpamFilterPolicy @OutboundParams
    $null = New-HostedOutboundSpamFilterRule -Name $OutboundSpamName -HostedOutboundSpamFilterPolicy$OutboundSpamName -SenderDomainIs $AcceptedDomains -Enabled$false
    Write-Host "Created Outbound Anti-Spam Policy: $OutboundSpamName" -ForegroundColor Green
} else {
    Write-Host "Outbound Anti-Spam Policy already exists: $OutboundSpamName" -ForegroundColor Yellow
}

Write-Host "`nIMPORTANT: All newly created policies are currently disabled and require manual enablement.`n" -ForegroundColor Yellow
