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
    Total  -> This field empty and used to add Size & Deleted together in this field inside Excel to determine the total size of a mailbox.
    SendAs 
    FullAccess


If you're having any issues with the script, please reach out to me.
https://github.com/ITR-MITHO

#>

Add-PSSnapin *EXC*
Import-Module ActiveDirectory

$File = Test-path "$Home\desktop\MailboxExport.csv"
If ($File)
{

$Confirm = Read-Host "MailboxExport.csv already present on your desktop. Do you want me to delete it for you? (Y/N)"
If ($Confirm -eq "Y")

{

Remove-Item "$Home\desktop\MailboxExport.csv" -Confirm:$false

}
    }


Clear-Host
$Mailboxes = Get-Mailbox -ResultSize Unlimited
$Results = @()
$StartDate = Get-Date


Write-Host "Start Time: $StartDate - It is estimated to take 10-15 minutes for large organisations. Grab a nice cup of coffee :-)" -ForegroundColor Yellow
Sleep 5
Foreach ($Mailbox in $Mailboxes)
{

$Statistics = Get-MailboxStatistics -Identity $Mailbox.SamAccountName
$Permission = Get-MailboxPermission -identity $Mailbox.SamAccountName | Where {$_.AccessRights -EQ "FullAccess" -and -not ($_.User -like “NT AUTHORITY\*”)}
$ADPermission = Get-Mailbox $Mailbox.SamAccountName | Get-ADPermission | Where {$_.ExtendedRights -like "Send-As" -and -not ($_.User -like “NT AUTHORITY\*”)}
$ADAtt = Get-ADUser -Identity $Mailbox.SamAccountName -Properties *

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
    Total = $null  # This field empty and used to add Size & Deleted together in this field inside Excel to determine the total size of a mailbox.
    SendAs = $ADPermission.User
    FullAccess = $Permission.User


}
    }
        
# Selecting the fields in a specific order instead of random.
$Results | Select Username, Name, Email, Type, Size, Deleted, Total, DB, LastLogon, ADEnabled, {$_.FullAccess}, {$_.SendAs} | 
Export-csv $home\Desktop\MailboxExport.csv -NoTypeInformation -Encoding Unicode

$EndDate = Get-Date
Write-Host "End Time: $EndDate - Find your .csv-file here: $Home\desktop\MailboxExport.csv" -ForegroundColor Green
