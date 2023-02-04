<#
V.2
The script will export the following information from all mailboxes: 
    SamAccountName
    DisplayName
    PrimarySmtpAddress
    RecipientTypeDetails
    DatabaseName
    LastLogonTime
    ADEnabled
    TotalItemSize.To.MB
    TotalDeletedItemSize.To.MB
    Total  "This field 0 and is used to add Size & Deleted together in Excel to determine the total size of a mailbox."
    SendAs 
    FullAccess


If you're having any issues with the script, please reach out to me.
https://github.com/ITR-MITHO

#>


Add-PSSnapin *EXC*
Import-Module ActiveDirectory

$Mailboxes = Get-Mailbox -ResultSize Unlimited
$Results = @()
$StartDate = Get-Date


Write-Host "The script is estimated to take about 10 minutes on larger envionments. Grab a coffee :-)" -ForegroundColor Yellow
Write-Host "Start Time $StartDate"
Sleep 5
Foreach ($Mailbox in $Mailboxes)
{
$MailboxUserName = $Mailbox.SamAccountName
$Statistics = Get-MailboxStatistics -Identity $MailboxUserName
$Permissions = Get-MailboxPermission -identity $MailboxUserName | Where {$_.AccessRights -EQ "FullAccess" -and -not ($_.User -like “NT AUTHORITY\*”)}
$ADPermission = Get-Mailbox $MailboxUserName | Get-ADPermission | Where {$_.ExtendedRights -like "Send-As" -and -not ($_.User -like “NT AUTHORITY\*”)}
$ADAtt = Get-ADUser -Identity $MailboxUserName -Properties *

if ($Statistics) 
{
    $Size = $Statistics.TotalItemSize.Value.ToMB()
    $Deleted = $Statistics.TotalDeletedItemSize.Value.ToMB()
} 
  
else 
{
    $Size = $null
    $Deleted = $null
}
Foreach ($Mailbox in $Mailboxes) {
$results += [PSCustomObject]@{

    Username = $Mailbox.SamAccountName
    Name = $Mailbox.DisplayName
    Email = $Mailbox.PrimarySmtpAddress
    Type = $Mailbox.RecipientTypeDetails
    DB = $Statistics.DatabaseName
    LastLogon = $Statistics.LastLogonTime
    ADEnabled = $ADAtt.Enabled
    Size = $Size
    Deleted = $Deleted
    Total = $null         # This field 0 and is used to add Size & Deleted together in Excel to determine the total size of a mailbox.
    SendAs = $ADPermission.User
    FullAccess = $Permission.User


}
    }
        }

# Selecting the fields in a specific order instead of random.
$Results | Select Username, Name, Email, Type, Size, Deleted, Total, DB, LastLogon, ADEnabled, {$_.FullAccess}, {$_.SendAs} | 
Export-csv $home\Desktop\MailboxExport.csv -NoTypeInformation -Encoding Unicode
$EndDate = Get-Date
Write-Host "Export completed! Find your .csv-file here: $Home\desktop\MailboxExport.csv" -ForegroundColor Green
Write-Host "End time $EndDate"
