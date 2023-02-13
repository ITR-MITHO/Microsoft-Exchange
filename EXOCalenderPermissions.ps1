<# 
 
.DESCRIPTION  
This script gives X (user or group) access to calendars for X users.
Default is all users, the accessrights can be changed under variables, so can the user/group.
 
 
.OUTPUTS 
It has a log, other than that, it just changes calendar permissions.
It's got a test function, if you remove the #/Hashtag from line 114 (the one that changes permissions)
 
 it outputs the file to "C:\office365Scripts\Keys\Credentials.txt"
If the file is not present, one will be made.
#>
 
 
#Variables
# Import the Exchange Online Management module
$UserName = "admin@onmicrosoft.com"
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

############################################# START OF ACTUAL SCRIPT #############################################

 
 
foreach($mbx in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox){
    $languageCalendar = (Get-MailboxFolderStatistics -Identity $mbx.userprincipalname -FolderScope Calendar | Select-Object -first 1).name
    Set-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User $UserToGiveAccess -AccessRights $AccessRight
    Get-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User $UserToGiveAccess | Select-Object identity,accessrights
}
