<#
.SYNOPSIS
    ITM8 Exchange Service - STANDARD
    Sets baseline/standard settings for an Exchange Online environment.
.DESCRIPTION
    - Enables Unified Audit Logging and ExternalInOutlook globally.
    - Sets Shared Mailbox sent items copy features.
    - Enforces Max Send/Receive sizes (150MB) and RetainDeletedItems (30 days).
    - Updates default calendar permissions to LimitedDetails.
#>

# ---------------------------------------------------------
# CONFIGURATION & PARAMETERS (Do Not Change Azure Vars)
# ---------------------------------------------------------
$ResourceGroup   = "rg-exchange"
$AutomationAccount = 'AA-Exchange'

# Customer Specific Information
$CertThumb        = 'CERTIFICATETHUMBPRINT'
$AppID            = 'APPID'
$OrganizationName = 'domain-com.onmicrosoft.com'

# Policy/Baseline Variables
$RetainDeletedItems = '30.00:00:00'
$MaxReceiveSize     = '150MB'
$MaxSendSize        = '150MB'
$ExternalInOutlook  = $true
$RetentionPolicy    = 'ITM8 - Deleted Items - 30 days'
$CalPer             = 'LimitedDetails'

# ---------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------
function Connect-EXOBaseline {
    [CmdletBinding()]
    param(
        [string]$Thumbprint,
        [string]$ApplicationId,
        [string]$Organization
    )
    process {
        Write-Verbose "Importing ExchangeOnlineManagement module..."
        Import-Module ExchangeOnlineManagement -ErrorAction Stop

        try {
            Connect-ExchangeOnline -CertificateThumbPrint $Thumbprint -AppID $ApplicationId -Organization $Organization -ErrorAction Stop
            Write-Output "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) - Connected to Exchange Online."
        }
        catch {
            Write-Error "Failed to connect to Exchange Online for $Organization. Error: $_"
            throw $_
        }
    }
}

function Optimize-MailboxCalendarPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Mailbox,
        [Parameter(Mandatory)][string]$AccessRights
    )
    process {
        try {
            $CalendarStats = Get-MailboxFolderStatistics -Identity $Mailbox.ExchangeGuid -FolderScope Calendar | Where-Object { $_.FolderType -eq 'Calendar' }
            foreach ($Folder in $CalendarStats) {
                $CalendarIdentity = "{0}:\{1}" -f $Mailbox.ExchangeGuid, $Folder.Name
                Set-MailboxFolderPermission -Identity $CalendarIdentity -User Default -AccessRights $AccessRights -WarningAction SilentlyContinue -ErrorAction Stop
            }
        }
        catch {
            Write-Warning "Failed to update calendar permissions for $($Mailbox.UserPrincipalName). Error: $_"
        }
    }
}

# ---------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------
$VerboseSettingOriginal = $VerbosePreference
$VerbosePreference = 'Continue' # Ensures Write-Verbose outputs to Azure Automation streams

Connect-EXOBaseline -Thumbprint $CertThumb -ApplicationId $AppID -Organization $OrganizationName

Write-Verbose "Configuring tenant-wide baseline settings..."
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -Confirm:$false
Set-ExternalInOutlook -Enabled $ExternalInOutlook

Write-Verbose "Fetching targeted mailboxes..."
$Mailboxes = Get-Mailbox -ResultSize Unlimited -Property ExchangeGuid, RecipientTypeDetails, UserPrincipalName, MaxSendSize, MaxReceiveSize, RetentionPolicy, RetainDeletedItemsFor

$TotalCount = $Mailboxes.Count
Write-Output "Processing settings across $TotalCount mailboxes..."

foreach ($M in $Mailboxes) {
    
    # Base parameters applicable to all mailboxes needing configuration updates
    $MailboxParams = @{
        Identity               = $M.ExchangeGuid
        MaxSendSize            = $MaxSendSize
        MaxReceiveSize         = $MaxReceiveSize
        RetentionPolicy        = $RetentionPolicy
        RetainDeletedItemsFor  = $RetainDeletedItems
        WarningAction          = 'SilentlyContinue'
    }
    if ($M.RecipientTypeDetails -eq 'SharedMailbox') {
        $MailboxParams['MessageCopyForSendOnBehalfEnabled'] = $true
        $MailboxParams['MessageCopyForSentAsEnabled']       = $true
    }
    try {
        Set-Mailbox @MailboxParams -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to update mailbox settings for $($M.UserPrincipalName). Error: $_"
    }

    # Update Calendar Permissions using the helper function
    Optimize-MailboxCalendarPermissions -Mailbox $M -AccessRights $CalPer
}

Write-Output "$(Get-Date) - Baseline optimization script completed. $TotalCount mailboxes evaluated."
$VerbosePreference = $VerboseSettingOriginal
