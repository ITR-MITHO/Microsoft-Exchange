# Deletes IIS logs that is older then 14-days
Import-Module WebAdministration
$Date = (Get-Date).AddDays(-14)
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
If ($Folder.Value -like "%Systemdrive%*")
{
    Get-ChildItem C:\Inetpub\Logs\LogFiles -Recurse | Where-Object {$_.LastWriteTime -LT $Date -and $_.Name -like "*.log"} | Remove-Item -Force
}
 Else
{   
    Get-ChildItem $Folder.Value -Recurse | Where-Object {$_.LastWriteTime -LT $Date} | Remove-Item -Force
}

# Deletes HTTPProxy logs that is older than 5 days
$ProxyDate = (Get-Date).AddDays(-5)
$HTTPProxy = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy"
Get-ChildItem $HTTPProxy -Recurse | Where-Object {$_.LastWriteTime -LT $ProxyDate -and $_.Name -like "*.log"} | Remove-Item -Force

# Deletes Exchange service logs that is older than 10 days
$ExchangeDate = (Get-Date).AddDays(-10)
Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Ews" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.log"} | Remove-Item -Force
Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Autodiscover" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\MapiHttp" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
