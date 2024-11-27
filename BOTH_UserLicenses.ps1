<#
Run fron on-prem Exchange. 
The script will prompt for O365 credentials to connect to MSOnline to gather license information about all on-prem mailboxes
If MSOnline module is missing, it will be installed

#>

<#

The script will prompt for O365 credentials to connect to MSOnline to gather license information about all on-prem mailboxes

#>

Add-PSSnapin *EXC*
Get-Mailbox -ResultSize unlimited | Select DisplayName, SamAccountName, DisplayName, UserPrincipalName, PrimarySMTPAddress, RecipientTypeDetails | Export-csv $home\desktop\LicenseCheck.csv -NoTypeInformation -Encoding Unicode

Try
{
Connect-MsolService -ErrorAction Stop
}
Catch
{

Write-Host "Installing the missing PowerShell Module: MSOnline. Please re-run the script afterwards" -ForegroundColor Yellow
Install-Module MSOnline -Confirm:$false
Break
}

$Mailboxes = Import-Csv $home\desktop\LicenseCheck.csv
$Results = @()

Foreach ($Mailbox in $Mailboxes) {
    $MailboxUPN = $Mailbox.UserPrincipalName
    $MailboxMail = $Mailbox.PrimarySMTPAddress
    $MailboxType = $Mailbox.RecipientTypeDetails
    $MailboxUsername = $Mailbox.SamAccountName
    $MailboxDisplayName = $Mailbox.DisplayName


    $License = (Get-MsolUser -UserPrincipalName $MailboxUPN -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Licenses -ErrorAction SilentlyContinue)
    If (-Not $License) {
        $License = "No license"
    } Else {

        $License = $License.AccountSkuId -join ", "
    }

    If ($License -like "*O365_w/o_Teams_Bundle_M5*") {
        $License = "Microsoft 365 E5 EEA (No teams)"
    } ElseIf ($License -like "*SPE_E3*") {
        $License = "Microsoft 365 E3"
    } Elseif ($License -like "*SPE_E5*") {
        $License = "Microsoft 365 E5"
    }

    $Results += [PSCustomObject]@{
        Username = $MailboxUsername
        Email    = $MailboxMail
        Licens   = $License
        Type     = $MailboxType
    }
}

$Results | Select-Object DisplayName, Username, Email, Licens, Type | Export-csv $home\desktop\Licenses.csv -NotypeInformation -Encoding Unicode
Write-Host "Export Completed, find your file here: $Home\Desktop\Licenses.csv" -ForeGroundColor Green
