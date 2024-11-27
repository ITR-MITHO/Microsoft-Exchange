<#
Run fron on-prem Exchange. 
The script will prompt for O365 credentials to connect to MSOnline to gather license information about all on-prem mailboxes
If MSOnline module is missing, it will be installed

#>

Add-PSSnapin *EXC*
Get-Mailbox -ResultSize unlimited | Select SamAccountName, DisplayName, PrimarySMTPAddress, RecipientTypeDetails | Export-csv $home\desktop\LicenseCheck.csv -NoTypeInformation -Encoding Unicode

Try
{
Connect-MsolService -ErrorAction Stop
}
Catch
{

Write-Host "Installing the missing PowerShell Module: MSOnline" -ForegroundColor Yellow
Install-Module MSOnline -Confirm:$false
Break
}

$Mailboxes = Import-csv $home\desktop\LicenseCheck.csv
$Results = @()
Foreach ($Mailbox in $Mailboxes)
{
$MailboxUPN = $Mailbox.PrimarySmtpAddress
$MailboxType = $Mailbox.RecipientTypeDetails
$MailboxUsername = $Mailbox.SamAccountName
$License = (Get-MsolUser -All | Where {$_.UserPrincipalName -eq $MailboxUPN} | Select @{Name='Licenses';Expression={$_.Licenses.AccountSkuId}})
$Statistics = Get-MailboxStatistics -Identity $MailboxUPN | Select-Object TotalItemSize

If (-Not $License)
{
$License = "No license"
}

If ($License -like "@*reseller-account:O365_w/o_Teams_Bundle_M5*")
{
$License = "Microsoft 365 E5 EEA (No teams)"
}

If ($Statistics) 
{
    $Size = $Statistics.TotalItemSize.Value.ToMB()
} 
Else 
{
    $Size = "0"
}


$Results += [PSCustomObject]@{
Username = $MailboxUsername
Email = $MailboxUPN
Licens = $License
Type = $MailboxType
SizeInMB = $Size
}
    }
$Results | Select-Object Username, Email, Licens, Type, Size
