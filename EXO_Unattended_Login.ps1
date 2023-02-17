<# 
 
.DESCRIPTION  
First time running the script, it will prompt you for the password to your O365 administrator.
Next time you run the script it will automatically use your username and password. 
The password is stored in a encrypted file inside C:\ITR

.NOTES
The only thing you need to change is the variable $Username to your O365 Administrator
To run use this in a scheduled script it is important to exclude the server or user you're using from Condtional Access/MFA policy.
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
$ITR = Test-Path "C:\ITR"
If (-Not $ITR)
{

mkdir "C:\ITR"

}

$PWFile = Test-Path "C:\ITR\Cred.txt"
if (-not $PWFile)
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
