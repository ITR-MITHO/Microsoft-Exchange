<# 
 
.DESCRIPTION  
First time running the script, it will prompt you for the credentials to your O365 administrator.
Next time you run the script it will automatically use your username and password. 

.NOTES
To use this in a scheduled script it is important to exclude the server or user you're using from Condtional Access/MFA policy.

#>

# Import the Exchange Online Management module
Try
{
Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
Catch
{
Write-Host "Exchange Online Module Missing" -ForegroundColor Yellow
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
Write-Host "Failed to connect to Exchange Online. Try to run 'Connect-ExchangeOnline' manually" -ForeGroundColor Red
Break
}
Write-host "Connected to Exchange Online!" -ForeGroundColor Green
