<# 

.DESCRIPTION  
Used to see what mailboxes are using EWS and which aren't using it.

.OUTPUT
1 file named "Activity.txt" will be placed on your desktop

#>

# Prerequisites
Add-PSSnapin *EXC*
$Date = (Get-Date).AddDays(-14)
$CSV = Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox, Sharedmailbox | Select SamAccountName, DisplayName, PrimarySMTPAddress, LastLogonDate
CD "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\Ews"
$Data = Get-ChildItem -Recurse | Where {$_.LastWriteTime -gt $Date}

cls
Write-Host "Starting to check logs from $Date til now" -ForeGroundColor Yellow
# Creating our own CSV-file with data
Echo "Name, Username, Email, LastLogon, Activity" | Out-File $home\desktop\Activity.txt


# Beginning to go through all mailboxes.
ForEach ($M in $csv)
{
$AD = Get-ADUser $M.SamAccountName -Properties LastLogonDate | Select LastLogonDate
$Name = $M.SamAccountName
$Full = $M.DisplayName
$Primary = $M.PrimarySMTPAddress
$Logon = $AD.LastLogonDate
$UserActivity = $Data | Select-String "$Name"
If (-not $UserActivity)
{

Echo "$Full, $Name, $Primary, $Logon, No" | Out-File $home\desktop\Activity.csv -Append
}
Else
{
Echo "$Full, $Name, $Primary, $Logon, Yes" | Out-File $home\desktop\Activity.csv -Append
}
 }
 cls
 Write-Host "Completed. Find your file here: $home\desktop\Activity.txt" -ForeGroundColor Green
