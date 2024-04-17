# Deletes IIS and HTTPProxy logs that is older then 14-days
Import-Module WebAdministration
$Date = (Get-Date).AddDays(-14)
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
If ($Folder.Value -like "%Systemdrive%*")
{
    Get-ChildItem C:\Inetpub\Logs\LogFiles -Recurse | Where-Object {$_.LastWriteTime -LT $Date} | Remove-Item -Force
}
 Else
{   
    Get-ChildItem $Folder.Value -Recurse | Where-Object {$_.LastWriteTime -LT $Date} | Remove-Item -Force
}

$HTTPProxy = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy"
Get-ChildItem $HTTPProxy -Recurse | Where-Object {$_.LastWriteTime -LT $Date} | Remove-Item -Force
