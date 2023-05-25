<#
.DESCRIPTION
The script will search IIS logs for the last 14 days to see which mailboxes are active and which aren't. 
If will place a .csv-file on your desktop named "Activity.csv" 

.NOTES
* Run PowerShell as administrator. 
* No editing is needed. 
#>

Add-PSSnapin *EXC*
Import-Module WebAdministration
Import-Module ActiveDirectory

$StartTime = Get-Date
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
$Date = (Get-Date).AddDays(-14)
$Mailbox = Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox, Sharedmailbox | Select SamAccountName, DisplayName, PrimarySMTPAddress, LastLogonDate

# Creating our own CSV-file with data
Echo "Name, Username, Email, LastLogon, Activity" | Out-File $home\desktop\Activity.csv -Encoding unicode

# Beginning to go through all mailboxes.
Clear

Write-Host "Starting to collect logs.. This might take a while." -ForegroundColor Yellow
ForEach ($M in $Mailbox)
{
$AD = Get-ADUser $M.SamAccountName -Properties LastLogonDate | Select LastLogonDate
$Name = $M.SamAccountName
$Full = $M.DisplayName
$Primary = $M.PrimarySMTPAddress
$Logon = $AD.LastLogonDate

If ($Folder.Value -like "%Systemdrive%*")
{
    CD "C:\Inetpub\Logs\LogFiles"
   $Data = Get-ChildItem -Recurse | Where {$_.LastWriteTime -GT $Date} | Select-String -Pattern "$Name" | Select -First 1
}
 Else
{  
    CD $Folder.Value
    $Data = Get-ChildItem -Recurse | Where {$_.LastWriteTime -GT $Date} | Select-String -Pattern "$Name" | Select -First 1
}

If (-not $Data)
{
Echo "$Full, $Name, $Primary, $Logon, No" | Out-File $home\desktop\Activity.csv -Append -Encoding unicode
}
Else
{
Echo "$Full, $Name, $Primary, $Logon, Yes" | Out-File $home\desktop\Activity.csv -Append -Encoding unicode
}
 }
 $EndTime = Get-Date
 Clear

 Write-Host "Completed. Find your file here: $home\desktop\Activity.txt" -ForeGroundColor Green
 Write-Host "Started at: $StartTime"
 Write-Host "Ended at: $EndTime"
