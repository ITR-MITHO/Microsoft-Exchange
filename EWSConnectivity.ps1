<# 

.DESCRIPTION  
Used to see what mailboxes are using EWS and which aren't using it.

.OUTPUT
2 files placed on desktop: YES-Activity and NO-Activity. 

#>

# Prerequisites
$Date = (Get-Date).AddDays(-30)
$CSV = Get-Mailbox -ResultSize unlimited | Select SamAccountName, DisplayName, PrimarySMTPAddress
CD "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\Ews"

# Creating our own CSV-file with data
Echo "Name, Username, Email" | Out-File $home\desktop\YES-activity.txt
Echo "Name, Username, Email" | Out-File $home\desktop\NO-activity.txt

# Beginning to go through all mailboxes.
ForEach ($M in $csv)
{
$Name = $M.SamAccountName
$Full = $M.DisplayName
$Primary = $M.PrimarySMTPAddress
$UserActivity = Get-ChildItem -Recurse | Where {$_.LastWriteTime -gt $Date} | Select-String "$Name" | Select LineNumber, Path -last 1
$Path = $UserActivity.Path
If (-not $UserActivity)
{

Echo "$Full, $Name, $Primary" | Out-File $home\desktop\NO-activity.txt -Append
}
Else
{
Echo "$Full, $Name, $Primary" | Out-File $home\desktop\YES-activity.txt -Append
}
 }
