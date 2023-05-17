<# 

.DESCRIPTION  
Used to see what mailboxes are using EWS and which aren't using it.

.OUTPUT
1 file named "Activity.txt" will be placed on your desktop

#>

# Prerequisites
$Date = (Get-Date).AddDays(-30)
$CSV = Get-Mailbox -ResultSize unlimited | Select SamAccountName, DisplayName, PrimarySMTPAddress, LastLogonDate
CD "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\Ews"

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
$UserActivity = Get-ChildItem -Recurse | Where {$_.LastWriteTime -gt $Date} | Select-String "$Name" | Select LineNumber, Path -last 1
$Path = $UserActivity.Path
If (-not $UserActivity)
{

Echo "$Full, $Name, $Primary, $Logon, No" | Out-File $home\desktop\Activity.txt -Append
}
Else
{
Echo "$Full, $Name, $Primary, $Logon, Yes" | Out-File $home\desktop\Activity.txt -Append
}
 }
