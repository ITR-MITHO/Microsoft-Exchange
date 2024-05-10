<#
Deletes logs in the following directories: 

LOGS OLDER THAN 14 DAYS
inetpub\logs\LogFiles

LOGS OLDER THAN 5 DAYS
$($env:ExchangeInstallPath)Logging\HttpProxy

LOGS OLDER THAN 10 DAYS
$($env:ExchangeInstallPath)Logging\Autodiscover
$($env:ExchangeInstallPath)Logging\EWS
$($env:ExchangeInstallPath)Logging\IMAP
$($env:ExchangeInstallPath)Logging\MAPI
$($env:ExchangeInstallPath)Logging\POP3

#>
Import-Module WebAdministration
$IISDays = (Get-Date).AddDays(-14)
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
If ($Folder.Value -like "%Systemdrive%*")
{
    Get-ChildItem C:\Inetpub\Logs\LogFiles -Recurse | Where-Object {$_.LastWriteTime -LT $IISDays -and $_.Name -like "*.LOG"} | Remove-Item -Force
}
 Else
{   
    Get-ChildItem $Folder.Value -Recurse | Where-Object {$_.LastWriteTime -LT $IISDays -and $_.Name -like "*.LOG"} | Remove-Item -Force
}

# Deletes HTTPProxy logs that is older than 5 days
$ProxyDate = (Get-Date).AddDays(-5)
Get-ChildItem "$($env:ExchangeInstallPath)Logging\HttpProxy" -Recurse | Where-Object {$_.LastWriteTime -LT $ProxyDate -and $_.Name -like "*.LOG"} | Remove-Item -Force

# Deletes Exchange service logs that is older than 10 days
$ExchangeDate = (Get-Date).AddDays(-10)
Get-ChildItem "$($env:ExchangeInstallPath)Logging\Ews" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
Get-ChildItem "$($env:ExchangeInstallPath)Logging\Autodiscover" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
Get-ChildItem "$($env:ExchangeInstallPath)Logging\MapiHttp" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
If (Test-Path "$($env:ExchangeInstallPath)Logging\Imap4")
{
Get-ChildItem "$($env:ExchangeInstallPath)Logging\Imap4" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
}
If (Test-Path "$($env:ExchangeInstallPath)Logging\POP3")
{
Get-ChildItem "$($env:ExchangeInstallPath)Logging\POP3" -Recurse | Where-Object {$_.LastWriteTime -LT $ExchangeDate -and $_.Name -like "*.LOG"} | Remove-Item -Force
}
