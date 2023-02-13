<# 
 
.DESCRIPTION  
This script gives everyone reviewer permissions to eachothers calendar. 

.NOTES
Change $UserToGiveAccess or $AccessRight if you want to modify who and what permission the script gives.
Comment out or remove line 26 after the first successful run to schedule the script to run unattended.
#>


# Import the Exchange Online Management module
$Username = "admin@domain.onmicrosoft.com"
Try
{
Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
Catch
{
Write-Host "Exchange Online Management Module is missing!" -ForegroundColor Red 
Write-Host "Install the missing module with PowerShell: Install-Module ExchangeOnlineManagement" -ForegroundColor Yellow
Break
}

# Store the credentials for the connection in a secure file
$cred = Get-Credential -UserName $UserName -Message "Enter password"
$cred.Password | ConvertFrom-SecureString | Set-Content -Path "C:\Office365\Keys\Cred.txt"

# Load the stored credentials
$securePassword = Get-Content -Path "C:\Office365\Keys\Cred.txt" | ConvertTo-SecureString
$cred = New-Object System.Management.Automation.PSCredential("$Username", $securePassword)

# Connect to Exchange Online
Connect-ExchangeOnline -Credential $cred -ShowProgress $true

# Check the connection status
if (Get-ExoSession) {
    Write-Host "Successfully connected to Exchange Online."
}
else {
    Write-Host "Failed to connect to Exchange Online."
}


# Reviewer Permissions for everyone to UserMailboxes
$UserToGiveAccess = 'Default' # Default is 'Everyone' in the organisation.
$AccessRight = 'Reviewer'
foreach ($mbx in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{

    $languageCalendar = (Get-MailboxFolderStatistics -Identity $mbx.userprincipalname -FolderScope Calendar | Select-Object -first 1).name
    Set-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User $UserToGiveAccess -AccessRights $AccessRight

}
