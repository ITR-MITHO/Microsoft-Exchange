<#
.DESCRIPTION
Works on Exchange 2013, 2016 and 2019. 
The script will search IIS logs for the last 5 days to see which mailboxes are active and which aren't.
Foeach Mailbox there is, it will search for that specific mailbox in all IIS-logs.
It will place a .csv-file on your desktop named "Activity.csv"

.NOTES
* Run PowerShell as administrator. 
* No editing is needed. 
* Can take 10-15 minutes to complete the exports. 

#>

# Checking permissions
$PMError = Test-Path $Home\desktop\PermissionIssue.txt
if ($PMError)
{
Remove-Item "$Home\desktop\PermissionIssue.txt" -Force
}
timeout 3 | Out-Null
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
echo "Start PowerShell as an Administrator" > $Home\desktop\PermissionIssue.txt
Start $home\desktop\PermissionIssue.txt
Break
}

# Modules
Add-PSSnapin *EXC*
Import-Module WebAdministration

# Variables
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
$Date = (Get-Date).AddDays(-5)
$Mailbox = Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox, Sharedmailbox | Select SamAccountName, DisplayName, PrimarySMTPAddress, LastLogonDate, RecipientTypeDetails

# Creating our own CSV-file with data
Echo "Name, Username, Email, LastLogon, Type, Activity" | Out-File $home\desktop\Activity.csv -Encoding unicode

# Beginning to go through all mailboxes.
Clear
Write-Host "Collecting and searching IIS-logs. This might take 10-15 minutes. 
When the report is done you will be notified in this box." -ForegroundColor Yellow

ForEach ($M in $Mailbox)
{
$MailboxStats = $M.SamAccountName | Get-MailboxStatistics | Select LastLogonTime -ErrorAction SilentlyContinue
if ($MailboxStats) 
{
    $LogonStat = $MailboxStats.LastLogonTime
} 
else
{
    $LogonStat = "No logon data"
}

$Name = $M.SamAccountName
$Full = $M.DisplayName
$Primary = $M.PrimarySMTPAddress
$Type = $M.RecipientTypeDetails
$Logon = $LogonStat

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
Echo "$Full, $Name, $Primary, $Logon, $Type, No" | Out-File $home\desktop\Activity.csv -Append -Encoding unicode
}
Else
{
Echo "$Full, $Name, $Primary, $Logon, $Type, Yes" | Out-File $home\desktop\Activity.csv -Append -Encoding unicode
}
 }
 
Clear
Write-Host "Completed. Find your file here: $home\desktop\Activity.csv" -ForeGroundColor Green
