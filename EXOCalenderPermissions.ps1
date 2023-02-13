<# 
 
.DESCRIPTION  
This script gives everyone reviewer permissions to eachother calendar. 
 
#>


# Import the Exchange Online Management module
$Username = "admin@onmicrosoft.com"
Import-Module ExchangeOnlineManagement

# Store the credentials for the connection in a secure file
$cred = Get-Credential -UserName $UserName -Message "Enter password" # Comment this line out after first run
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
