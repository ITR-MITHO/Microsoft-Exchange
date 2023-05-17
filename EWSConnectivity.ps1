<# 

.DESCRIPTION  
Used to see what mailboxes are using EWS and which aren't using it.

.OUTPUT
2 files placed on desktop: YES-Activity and NO-Activity. 

#>

$Date = (Get-Date).AddDays(-30)
$CSV = Get-Mailbox -ResultSize unlimited | Select SamAccountName, DisplayName
CD "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\Ews"
Echo "Name, Username, Path" Out-File $home\desktop\YES-activity.txt

ForEach ($M in $csv)
{
$Name = $M.SamAccountName
$Full = $M.DisplayName
$UserActivity = Get-ChildItem -Recurse | Where {$_.LastWriteTime -gt $Date} | Select-String "$Name" | Select LineNumber, Path -last 1
$Path = $UserActivity.Path
If (-not $UserActivity)
{

Echo "$Full - $Name - have NO activity" | Out-File $home\desktop\NO-activity.txt -Append
}
Else
{
Echo "$Full, $Name, $Path" | Out-File $home\desktop\YES-activity.txt -Append
}
 }
