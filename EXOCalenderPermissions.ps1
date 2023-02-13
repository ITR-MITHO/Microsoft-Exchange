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
Try
{
Connect-ExchangeOnline -Credential $cred -ShowProgress $true -ErrorAction Stop
}
Catch
{
Write-Host "Connect-ExchangeOnline failed" -ForeGroundColor Red
Break
}

# Check the connection status
Try
{
Get-ExoSession -ErrorAction Stop
}
Catch
{
Write-Host "Failed to connect to Exchange Online." -ForeGroundColor Red
Break
}
Write-host "Connected to Exchange Online!" -ForeGroundColor Green

# Reviewer Permissions for everyone to UserMailboxes
# Default UserMailbox Calendar Permissions
$User = 'Default'
$AccessRight = 'Reviewer'
Foreach ($Mailbox in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{

    $Calendar = (Get-MailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -FolderScope Calendar | Select-Object -First 1).Name
    Set-MailboxFolderPermission -Identity ($Mailbox.UserPrincipalName+":\$Calendar") -User $User -AccessRights $AccessRight

}
