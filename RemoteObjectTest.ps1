# Check elevated state of Powershell
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
write-host "Script is not running as Administrator" -ForegroundColor Yellow
Break
}

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

# Connecting to ExchangeOnline
Try
{
Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
}
Catch
{
Write-Host "Failed to connect to Exchange Online." -ForegroundColor Red    
Write-Host "Try to run 'Connect-ExchangeOnline' manually" -ForeGroundColor Yellow
Break
}
Write-host "Connected to Exchange Online!" -ForeGroundColor Green

# Gather Mailbox Information
$EXOMailbox = Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails UserMailbox | Select-Object Alias, PrimarySmtpAddress, @{Name="EmailAddresses";Expression={($_.EmailAddresses | Where-Object {$_ -clike "smtp*"} | ForEach-Object {"`"$_`","}) -join " "}}
$Results = @()
Foreach ($EXO in $EXOMailbox)
    {


    $results += [PSCustomObject]@{
        Alias = $EXO.Alias
        PrimarySMTPAddress = $EXO.PrimarySMTPAddress
        EmailAddresses = $EXO.EmailAddresses

}
    }

$Results | Select Alias, PrimarySMTPAddress, EmailAddresses | Export-csv $home\desktop\EXOMailboxes.csv -NoTypeInformation -Encoding Unicode

# Disconnecting from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Add-PSSnapin *EXC*
$CSV = Import-csv $Home\desktop\EXOMailboxes.csv
Echo "Alias, PrimarySMTPAddress, Emailaddresses" > $home\desktop\missing.csv
Foreach ($C in $CSV)
{

$Alias = $C.Alias
$Primary = $C.PrimarySMTPAddress
$Email = $C.EmailAddresses

If (Get-Recipient $C.Alias -ErrorAction SilentlyContinue)
{
 
}
Else
{
Echo "$Alias, $Primary, $Email" | Out-File $home\desktop\Missing.csv -Append
}
    }  

Write-Host "Objects missing remote routing objects can be found in $home\desktop\missing.csv" -ForegroundColor Yellow