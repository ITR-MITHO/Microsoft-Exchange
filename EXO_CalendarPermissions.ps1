<# 
 
.DESCRIPTION  
This script gives everyone reviewer permissions to eachothers calendar.


.NOTES
Change the variable $Username to your tenant administrator name.
Change the variable $User or $AccessRight if you want to modify who and what permission the script gives.

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
$TestPath = Test-Path "C:\ITR\Cred.txt"
if (-not $TestPath)
{

    $cred = Get-Credential -UserName $UserName -Message "Enter password"
    $cred.Password | ConvertFrom-SecureString | Set-Content -Path "C:\ITR\Cred.txt"

}

# Load the stored credentials
$securePassword = Get-Content -Path "C:\ITR\Cred.txt" | ConvertTo-SecureString
$cred = New-Object System.Management.Automation.PSCredential("$Username", $securePassword)

# Connect to Exchange Online
Try
{
Connect-ExchangeOnline -Credential $cred -ShowProgress $true -ErrorAction Stop
}
Catch
{
Write-Host "Failed to connect to Exchange Online. Try to run 'Connect-ExchangeOnline' manually" -ForeGroundColor Red
Break
}
Write-host "Connected to Exchange Online!" -ForeGroundColor Green


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
