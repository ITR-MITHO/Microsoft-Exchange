   <# 
   
   The script will collect the 500 newest entries from a specific user in the IIS logs the past X days on the following services: 
   Autodiscover, Exchange Web Services and MAPI
   Change the variable $Days to something else if you want it to be higher than 3 days.
   The only thing you need to do, is to run this script as elevated, and enter a username or IP-address to search for.
  
   #>

Import-Module WebAdministration
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
$Days = 3
$PMError = Test-Path $Home\desktop\ExchangeLogs
if ($PMError)
{

}
Else
{
mkdir $home\desktop\ExchangeLogs | Out-null
}
   $User = Read-Host "Enter the username of the account you'd like to search for or the IP-address you are interrested in"

   If ($Folder.Value -like "%Systemdrive%*")
{
    CD "C:\Inetpub\Logs\LogFiles\W3SVC1"
}
 Else
{  
    CD $Folder.Value
}
CLS
Write-Host "INFORMATION: Searching for $User - This can take 2-5 minutes." -foregroundcolor Yellow
$Data = Get-ChildItem -Recurse | Sort-Object -Descending | Select -First $Days

$Data | Select-String -Pattern "$User" | Where {$_.Line -like "*/MAPI/*"} | Sort-Object -Descending | Select -first 500 > $home\Desktop\ExchangeLogs\MAPI.log
$Data | Select-String -Pattern "$User" | Where {$_.Line -like "*/Autodiscover/*"} | Sort-Object -Descending | Select -first 500 > $home\Desktop\ExchangeLogs\Autodiscover.log
$Data | Select-String -Pattern "$User" | Where {$_.Line -like "*/EWS/*"} | Sort-Object -Descending | Select -first 500 > $home\Desktop\ExchangeLogs\EWS.log

CLS
Write-Host "INFORMATION: Completed! Find your files $Home\Desktop\Exchangelogs" -ForegroundColor Green
