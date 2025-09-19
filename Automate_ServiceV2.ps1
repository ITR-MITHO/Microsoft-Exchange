<#

ITM8 Exchange Service - STANDARD
This script sets our baseline/standard settings for a Exchange Online environment

ExternalInOutlook - Is set for everyone
LimitedDetails - Is set as the default calendar permission on UserMailboxes
RoomMailbox - A set of default parameters
SharedMailbox - Save sent items in the mailbox
Backup - Every 15 day a export of all mailboxes attributes is exported

#>

$ErrorActionPreference = 'SilentlyContinue'
$Customer = Import-csv "C:\ITM8\Customers.csv"
Foreach ($C in $Customer)
{

# Veriables
$Mailbox = Get-Mailbox -RecipientTypeDetails UserMailbox
$Count = ($Mailbox.Count)
$Date = Get-Date
$OrgName = $C.Org

# Connect to Exchange Online
Connect-ExchangeOnline -AppID $C.App -CertificateThumbprint $C.Thumb -Organization $C.Org

# MailTips
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
Set-ExternalInOutlook -Enabled $true

# Shared Mailbox - Save sent items in the mailbox
Get-Mailbox -RecipientTypeDetails SharedMailbox | Set-Mailbox -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true

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


<# Default RoomMailbox Calendar Processing
$Parameter = @{
AutomateProcessing = "AutoAccept"
DeleteComments = $true
AddOrganizerToSubject = $true
AllowConflicts = $false
ProcessExternalMeetingMessages = $false
BookingWindowInDays = "180"
MaximumDurationInMinutes = "600"
MinimumDurationInMinutes = "5"
}
Foreach ($Room in Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails RoomMailbox)
{
Try
{
    $UserPrincipalName = $Room.UserPrincipalName
    Set-CalendarProcessing -identity $UserPrincipalName @Parameter
}
Catch
{
    Write-Warning "Failed to update CalendarProcessing on $UserPrincipalName"
    Continue
}
    }

#>
# Simple logging

Echo "$DATE - changes made to $Count mailboxes" >> "C:\ITM8\$OrgName\$OrgName.log"

# Close ExchangeOnline Session before starting a export
Disconnect-ExchangeOnline -Confirm:$false

# Re-connecting
Connect-ExchangeOnline -AppID $C.App -CertificateThumbprint $C.Thumb -Organization $C.Org
$File = Get-ChildItem -Path "C:\ITM8\$OrgName\Backup.csv"
If ($File.LastModified -LT (Get-Date).AddDays(-15))
{
Get-Mailbox | Select * | Export-csv C:\ITM8\$OrgName\$Date-Backup.csv -NotypeInformation -Encoding UNICODE
Disconnect-ExchangeOnline -Confirm:$false
}
Else
{
Disconnect-ExchangeOnline -Confirm:$false
}

    }
