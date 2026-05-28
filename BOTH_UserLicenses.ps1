<#
Run fron on-prem Exchange. 
The script will prompt for O365 credentials to connect to Microsoft Graph to gather license information about all on-prem mailboxes
If MSOnline module is missing, it will be installed

#>

Add-PSSnapin *EXC*
Get-Mailbox -ResultSize unlimited | Select SamAccountName, DisplayName, UserPrincipalName, PrimarySMTPAddress, RecipientTypeDetails | Export-csv $home\desktop\LicenseCheck.csv -NoTypeInformation -Encoding Unicode

Try
{
Connect-MgGraph -Scopes User.Read.All, Organization.Read.All -ErrorAction Stop
}
Catch
{

Write-Host "Installing the missing PowerShell Module: Microsoft Graph. Please re-run the script afterwards
This can take 5-10 minutes..." -ForegroundColor Yellow
Install-Module Microsoft.Graph -Scope CurrentUser -Confirm:$false
Break
Write-Host "Microsoft Graph installed, re-run the script" -ForegroundColor Green
}

$Mailboxes = Import-Csv $home\desktop\LicenseCheck.csv
$Results = @()

Foreach ($Mailbox in $Mailboxes) {
    $MailboxUPN = $Mailbox.UserPrincipalName
    $MailboxMail = $Mailbox.PrimarySMTPAddress
    $MailboxType = $Mailbox.RecipientTypeDetails
    $MailboxUsername = $Mailbox.SamAccountName
    $MailboxDisplayName = $Mailbox.DisplayName
    $ADAtt = Get-ADUser -Identity $MailboxUserName -Properties Enabled, LastLogonDate


    $License = (Get-MgUserLicenseDetail -UserId $MailboxMail -ErrorAction SilentlyContinue | Select-Object SkuPartNumber -ErrorAction SilentlyContinue)
    If (-Not $License) {
        $License = "No license"
    } Else {

        $License = $License.SkuPartNumber -join ", "
    }

    If ($License -like "*SPE_E3*") {
        $License = "Microsoft 365 E3"

    } Elseif ($License -like "*SPE_E5*") {
        $License = "Microsoft 365 E5"

    } Elseif ($License -like "*SPB*") {
        $License = "Microsoft Business Premium"

    } Elseif ($License -like "*EXCHANGESTANDARD*") {
        $License = "Exchange Online Plan 1"

    } Elseif ($License -like "*EXCHANGEPREMIUM*") {
        $License = "Exchange Online Plan 2"
    }


    IF ($ADAtt.LastLogonDate)
{
$LastLogonDate = $ADAtt.LastlogonDate.ToString("dd-MM-yyyy")
}
Else
{
$LastLogonDate = ""
}


    $Results += [PSCustomObject]@{
        DisplayName = $MailboxDisplayName
        Username = $MailboxUsername
        Email    = $MailboxMail
        Licens   = $License
        Type     = $MailboxType
        Enabled  = $ADAtt.Enabled
        LastLogon = $LastLogonDate

    }
}

$Results | Select-Object DisplayName, Username, Email, Licens, Type, Enabled, LastLogon | Export-csv $home\desktop\Licenses.csv -NotypeInformation -Encoding Unicode -Delimiter ";"
Write-Host "Export Completed, find your file here: $Home\Desktop\Licenses.csv" -ForeGroundColor Green
