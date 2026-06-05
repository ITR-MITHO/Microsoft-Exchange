<#
.SYNOPSIS
    Standardizes default calendar permissions across all user mailboxes (Multilingual Safe).
.DESCRIPTION
    Iterates through all on-premises user mailboxes, dynamically discovers the localized 
    name of the Calendar folder (e.g., 'Calendar', 'Kalender'), and updates the 'Default' 
    user permission level to 'Reviewer'.
.NOTES
    Must be executed from an elevated Exchange Management Shell session.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

$TargetUser  = 'Default'
$AccessRight = 'Reviewer'

Write-Host "Gathering user mailboxes..." -ForegroundColor Cyan
$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
$MailboxCount = $Mailboxes.Count
$Count = 1

Write-Host "Applying calendar permissions..." -ForegroundColor Cyan
foreach ($Mailbox in $Mailboxes) {
    $Upn = $Mailbox.UserPrincipalName
    
    # Progress Bar UI Control
    $Activity = 'Processing... [{0}/{1}]' -f $Count, $MailboxCount
    $Status   = 'Updating calendar defaults for: {0}' -f $Upn
    Write-Progress -Status $Status -Activity $Activity -PercentComplete (($Count / $MailboxCount) * 100)

    try {
        # Dynamically resolve the localized name of the Calendar folder using the well-known folder type
        # This returns 'Calendar' for English, 'Kalender' for Danish, etc., instantly.
        $LocalizedCalendarName = (Get-MailboxFolder -Identity "$Upn:\Calendar" -ErrorAction Stop).Name
        $CalendarIdentity = "$Upn:\$LocalizedCalendarName"

        # Attempt to set the permission assuming the default entry exists
        Set-MailboxFolderPermission -Identity $CalendarIdentity -User $TargetUser -AccessRights $AccessRight -SendNotificationToUser $false -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "Success: Set $TargetUser to $AccessRight on $CalendarIdentity" -ForegroundColor Green
    } 
    catch {
        # Check if the specific exception means the permission entry doesn't exist yet
        if ($_.Exception.Message -match "ExistingMailboxFolderPermissionNotFoundException" -or $_.Exception.Message -match "no existing permission") {
            try {
                # Fallback to Add if the rule mapping does not exist yet
                Add-MailboxFolderPermission -Identity $CalendarIdentity -User $TargetUser -AccessRights $AccessRight -SendNotificationToUser $false -ErrorAction Stop
                Write-Host "Success: Added $TargetUser with $AccessRight on $CalendarIdentity" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to add calendar permission entry for $Upn. Reason: $($_.Exception.Message)"
            }
        } else {
            # Capture real errors like offline databases, connection limits, or missing folders
            Write-Warning "Unexpected exception processing $Upn. Reason: $($_.Exception.Message)"
        }
    }
    $Count++
}

# Clear the progress bar when complete
Write-Progress -Activity "Processing..." -Completed
Write-Host "`nCalendar permission standardization complete." -ForegroundColor Green
