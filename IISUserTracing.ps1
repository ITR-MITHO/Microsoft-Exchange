<# 
   
   The script will collect the newest entries from a specific user in the IIS logs from the past X days on the following services: 
   Autodiscover, Exchange Web Services, MAPI and ActiveSync

   Change the variable $Days to something else if you want it to be higher than 1 day.
   The only thing you need to do, is to run this script as elevated, and enter a username or IP-address to search for.
  
   #>

$Days = 1
Import-Module WebAdministration
$IISFolder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
$Date = (Get-Date).AddDays(-$Days)
$DesktopFolder = Test-Path $Home\desktop\ExchangeLogs
if (-not $Desktopfolder)
{
mkdir $home\desktop\ExchangeLogs | Out-null
}
   $User = Read-Host "Enter the username of the account you'd like to search for or the IP-address you are interrested in"
   
   If ($IISFolder.Value -like "%Systemdrive%*")
{
    CD "C:\Inetpub\Logs\LogFiles\W3SVC1"
}
 Else
{  
    CD $IISFolder.Value
}
CLS
Write-Host "INFORMATION: Searching for $User in IIS logs for the past $Days days." -foregroundcolor Yellow
$Data = Get-ChildItem -Recurse | Where {$_.LastWriteTime -GT $Date} | Sort-Object -Descending

$Data | Select-String -Pattern "$User" | Where {$_.Line -like "*/Autodiscover/*"} | Sort-Object -Descending > $home\Desktop\ExchangeLogs\Autodiscover.log
$Data | Select-String -Pattern "$User" | Where {$_.Line -like "*/EWS/*"} | Sort-Object -Descending > $home\Desktop\ExchangeLogs\EWS.log
$Data | Select-String -Pattern "$User" | Where {$_.Line -like "*/MAPI/*"} | Sort-Object -Descending > $home\Desktop\ExchangeLogs\MAPI.log
$Data | Select-String -Pattern "$User" | Where {$_.Line -like "*/Microsoft-Server-ActiveSync/*"} | Sort-Object -Descending > $home\Desktop\ExchangeLogs\ActiveSync.log


Write-Host "INFORMATION: Find your log files here: $Home\Desktop\Exchangelogs" -ForegroundColor Green
