Add-PSSnapin *EXC*
Import-Module WebAdministration

$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
$Date = (Get-Date).AddDays(-14)
$Mailbox = Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox, Sharedmailbox | Select SamAccountName, DisplayName, PrimarySMTPAddress, LastLogonDate

# Creating our own CSV-file with data
Echo "Name, Username, Email, LastLogon, Activity" | Out-File $home\desktop\Activity.csv

# Beginning to go through all mailboxes.
CLS
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
    # Change the line below with one of the following: */mapi/*, */Microsoft-Server-ActiveSync/*, */AutoDiscover/*, */OWA/*, */ECP/* and */EWS/*
   $Data = Get-ChildItem -Recurse | Where {$_.LastWriteTime -GT $Date} | Select-String -Pattern "$Name" | Where {$_.Line -like "*/EWS/*"} | Select -First 1
}
 Else
{  
    CD $Folder.Value
    # Change the line below with one of the following: */mapi/*, */Microsoft-Server-ActiveSync/*, */AutoDiscover/*, */OWA/*, */ECP/* and */EWS/*
    $Data = Get-ChildItem -Recurse | Where {$_.LastWriteTime -GT $Date} | Select-String -Pattern "$Name" | Where {$_.Line -like "*/EWS/*"} | Select -First 1
}


If (-not $Data)
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
