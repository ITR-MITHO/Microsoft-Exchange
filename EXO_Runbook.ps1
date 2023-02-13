# Fill out the below variables with it's static information
$ResourceGroupID = 'ResourceGroup ID'
$AutomationAccount = 'AutomationAccount ID'
$Certname = 'Certificate name'
$AppID = 'Application ID'
$OrganizationName = 'organization name e.g.: itr.onmicrosoft.com'

# Obtain a copy of the certificate
$ExchangeOnlineCertThumbPrint = (Get-AzAutomationCertificate -ResourceGroupName "$ResourceGroupID" -AutomationAccountName "$AutomationAccount" -Name "$Certname").Thumbprint

# Import Exchange Online Module
Try
{
Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
Catch
{
Write-Host "Exchange Online Management Module is missing!" -ForegroundColor Red 
Break
}

# Connect to Exchange Online using the Certificate Thumbprint of the Certificate imported into the Automation Account
Connect-ExchangeOnline -CertificateThumbPrint $ExchangeOnlineCertThumbPrint -AppID "$AppID" -Organization "$OrganizationName"


# Default Calendar Permissions

$UserToGiveAccess = 'Default'
$AccessRight = 'Reviewer'
foreach ($mbx in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{

    $languageCalendar = (Get-MailboxFolderStatistics -Identity $mbx.userprincipalname -FolderScope Calendar | Select-Object -first 1).name
    Set-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User $UserToGiveAccess -AccessRights $AccessRight

}


# Default Room Mailbox Configuration

$Processing = 'AutoAccept'
$DeleteComments = $true
$OrganizaerToSubject = $true
$Conflicts = $false
$ExternalMeetings = $false
Foreach ($Room in Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails RoomMailbox)
{

Set-CalendarProcessing -Identity $Room.Alias -AutomateProcessing $Processing -DeleteComments $DeleteComments -AddOrganizerToSubject $OrganizaerToSubject -AllowConflicts $Conflicts -ProcessExternalMeetingMessages $ExternalMeetings

}
