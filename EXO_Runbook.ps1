<#
.DESCRIPTION
The script is designed to use within Azure Runebooks. 
It will set user calendar permissions for all UserMailboxes and setup a default behaviour for RoomMailboxes. 

.NOTES
Edit the settings for UserMailbox calendar permissions in line 50 & 51
Edit the settings for RoomMailbox behaviour in line 70-78

#>

# Fill out the below variables with static information
$ResourceGroup = "MITHO-ResourceGroup"
$AutomationAccount = 'MITHO-AutomationAccount'
$Certname = 'mycert' # Name of the cert added to the automation account
$AppID = 'XXXXX-XXXXXX-XXXXXX-XXXXXXX'
$OrganizationName = 'itrmitho.onmicrosoft.com'

# Azure Context
Disable-AzContextAutosave -Scope Process
$AzureContext = (Connect-AzAccount -Identity).context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext

# Obtain a copy of the certificate
$ExchangeOnlineCertThumbPrint = (Get-AzAutomationCertificate -ResourceGroupName "$ResourceGroup" -AutomationAccountName "$AutomationAccount" -Name "$Certname").Thumbprint

# Import Exchange Online Module
Try
{
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
    Catch
{
    Write-Warning "Exchange Online Management Module is missing!"
    Break
}

# Connect to Exchange Online using the Certificate Thumbprint of the Certificate imported into the Automation Account
Try
{
    Connect-ExchangeOnline -CertificateThumbPrint $ExchangeOnlineCertThumbPrint -AppID "$AppID" -Organization "$OrganizationName" -ErrorAction Stop
}
Catch
{
Write-Warning "Failed to connect to Exchange Online. Ensure that the certificate is valid!"
Break
}

# Default UserMailbox Calendar Permissions
$User = 'Default'
$AccessRight = 'Reviewer'
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

# Default RoomMailbox Calendar Processing
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
