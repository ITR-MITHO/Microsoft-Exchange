<# 
   
   The script will collect the newest entries from a specific user in the IIS logs from the past X hours on the following services: 
   Autodiscover, Exchange Web Services, MAPI and ActiveSync

   Change the variable $Hours to something else if you want it to be higher than 3 hours.
   The only thing you need to do, is to run this script as elevated, and enter a username or IP-address to search for.
  
  #>

$Hours = 12 # Whatever number you type will be X before current time. So if it is 12 it will be 12 hours ago.
Import-Module WebAdministration
$IISFolder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
$Date = (Get-Date).AddHours(-$Hours)
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
Write-Host "INFORMATION: Searching for $User in logs..." -foregroundcolor Yellow
$Data = Get-ChildItem -Recurse | Where {$_.LastWriteTime -GT $Date} | Sort-Object -Descending

$Data | Select-String -Pattern "$User" | sls "/Autodiscover/" > $home\Desktop\ExchangeLogs\Autodiscover.log
$Data | Select-String -Pattern "$User" | sls "/EWS/" > $home\Desktop\ExchangeLogs\EWS.log
$Data | Select-String -Pattern "$User" | sls "/MAPI/" > $home\Desktop\ExchangeLogs\MAPI.log
$Data | Select-String -Pattern "$User" | sls "/OAB/" > $home\Desktop\ExchangeLogs\OAB.log
$Data | Select-String -Pattern "$User" | sls "/OWA/" > $home\Desktop\ExchangeLogs\OWA.log
$Data | Select-String -Pattern "$User" | sls "/ECP/" > $home\Desktop\ExchangeLogs\ECP.log
$Data | Select-String -Pattern "$User" | sls "/Microsoft-Server-ActiveSync/" > $home\Desktop\ExchangeLogs\ActiveSync.log

Write-Host "INFORMATION: Find your log files here: $Home\Desktop\Exchangelogs" -ForegroundColor Green
