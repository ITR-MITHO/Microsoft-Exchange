<# 
 
.DESCRIPTION 
Gives everyone LimitedDetails permissions to any calendar.
First time running the script, it will prompt you for the credentials to your O365 administrator.
Next time you run the script it will automatically use your username and password. 

.NOTES
To use this in a scheduled script it is important to exclude the server or user you're using from Condtional Access/MFA policy.
Edit line 62 & 63 to change user and accessrights.

#>

# Import the Exchange Online Management module
Try
{
Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
Catch
{
Write-Host "Exchange Online Module is NOT installed" -ForegroundColor Red
Write-Host "Install the missing module with PowerShell: Install-Module ExchangeOnlineManagement" -ForegroundColor Yellow
Break
}

# Store the credentials for the connection in a secure file
$ITR = Test-Path "C:\ITR\EXO"
If (-Not $ITR)
{
mkdir "C:\ITR\EXO"
}

$PWFile = Test-Path "C:\ITR\EXO\PWD.txt"
if (-not $PWFile)
{
    $cred = Get-Credential -Message "Enter your O365 administrator credentials 
(example: admin@domain.onmicrosoft.com)"
    $cred.Password | ConvertFrom-SecureString | Set-Content -Path "C:\ITR\EXO\PWD.txt"
    $Cred.UserName | Set-Content -path "C:\ITR\EXO\USR.txt"
}

# Load the stored credentials
$SecurePassword = Get-Content -Path "C:\ITR\EXO\PWD.txt" | ConvertTo-SecureString
$SecureUser = Get-Content -Path "C:\ITR\EXO\USR.txt"
$cred = New-Object System.Management.Automation.PSCredential($SecureUser, $securePassword)

# Connect to Exchange Online
Try
{
Connect-ExchangeOnline -Credential $cred -ShowProgress $true -ErrorAction Stop
}
Catch
{
Write-Host "Failed to connect to Exchange Online. If your user requires Multi-factor authentication from this destination, it will not work." -ForegroundColor Red    
Write-Host "Try to run 'Connect-ExchangeOnline' manually, to see if it prompts for MFA" -ForeGroundColor Yellow
Break
}
Write-host "Connected to Exchange Online!" -ForeGroundColor Green


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
    Write-Warning "Failed to add the permissions on Mailbox: $UserPrincipalName"
    Continue
}
    }
