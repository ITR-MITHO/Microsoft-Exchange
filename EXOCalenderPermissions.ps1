<# 
 
.DESCRIPTION  
This script gives everyone reviewer permissions to eachothers calendar.
It can be run manually prompting for a password or scheduled to run unattended.

.NOTES
Change $Username to your tenant administrator name.
Change $User or $AccessRight if you want to modify who and what permission the script gives.
Comment out or remove line 28 after the first successful run to schedule the script to run unattended.
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
Write-Host "Failed to connect to Exchange Online. Try to run the cmdlet manually" -ForeGroundColor Red
Break
}
Write-host "Connected to Exchange Online!" -ForeGroundColor Green


# Default UserMailbox Calendar Permissions
$User = 'Default'
$AccessRight = 'Reviewer'
Foreach ($Mailbox in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{
Try 
{
    $Calendar = (Get-MailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -FolderScope Calendar | Select-Object -First 1).Name
    Set-MailboxFolderPermission -Identity ($Mailbox.UserPrincipalName+":\$Calendar") -User $User -AccessRights $AccessRight -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
}
Catch
{
    Write-host "Failed to add the user '$User' with calendar permission '$AccessRight' on mailbox: $Mailbox" -ForeGroundColor Red
}
    Write-Host "Sucessfully added the user '$User' with calendar permissions '$AccessRight' on Mailbox: $Mailbox" -ForegroundColor Green
}
