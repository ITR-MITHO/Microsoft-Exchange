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
Write-Warning "Connect-ExchangeOnline failed. Ensure that the certificate is valid!"
Break
}


# Default UserMailbox Calendar Permissions
$User = 'Default'
$AccessRight = 'Reviewer'
Foreach ($Mailbox in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{
Try 
{
    $UserPrincipalName = $Mailbox.UserPrincipalName
    $Calendar = (Get-MailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -FolderScope Calendar | Select-Object -First 1).Name
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
