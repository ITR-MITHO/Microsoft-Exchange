<#
.DESCRIPTION
The script will prompt for Exchange Online credentials to export a list of all existing EXO mailboxes
Afterwards it will from on-prem Exchange check if it can see the mailboxes. If it can't see the mailbox, it indicates that the remote routing objects is missing and needs to be enabled. 

.SYNOPSIS
Use from on-prem Exchange server, with elevated shell

#>

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

# Gather Mailbox Information
$Domain = Read-Host "Enter Tenant name (e.g. contoso.mail.onmicrosoft.com)"
$EXOMailbox = Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails UserMailbox | Select-Object Alias, PrimarySmtpAddress, @{Name="EmailAddresses";Expression={($_.EmailAddresses | ForEach-Object {"`"$_`","}) -join ""}}
$Results = @()
Foreach ($EXO in $EXOMailbox)
    {
    $results += [PSCustomObject]@{
        Alias = $EXO.Alias
        PrimarySMTPAddress = $EXO.PrimarySMTPAddress
        EmailAddresses = $EXO.EmailAddresses
        RemoteRouting = "$($EXO.Alias)@$Domain"

}
    }

$Results | Select-Object Alias, PrimarySMTPAddress, RemoteRouting, EmailAddresses | Export-csv $home\desktop\EXOMailboxes.csv -NoTypeInformation -Encoding Unicode

# Disconnecting from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Add-PSSnapin *EXC*
$CSV = Import-csv $Home\desktop\EXOMailboxes.csv
Echo "Alias, PrimarySMTPAddress, RemoteRouting, Emailaddresses" > $home\desktop\RemoteMissing.csv
Foreach ($C in $CSV)
{

$Alias = $C.Alias
$Primary = $C.PrimarySMTPAddress
$Email = $C.EmailAddresses.TrimEnd(',')
$RemoteRouting = $C.RemoteRouting

If (Get-Recipient $C.Alias -ErrorAction SilentlyContinue)
{
}
Else
{
Echo "$Alias, $Primary, $RemoteRouting, $Email" | Out-File $home\desktop\RemoteMissing.csv -Append
}
    }  
Write-Host "Objects missing remote routing objects can be found in $home\desktop\RemoteMissing.csv" -ForegroundColor Green
Remove-Item $Home\Desktop\EXOMailboxes.csv -Force
