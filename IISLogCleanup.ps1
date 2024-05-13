<#
Deletes .LOG files older than 10 days in the following directories: 

inetpub\logs\LogFiles
$($env:ExchangeInstallPath)Logging\*

#>
Import-Module WebAdministration
$Days = (Get-Date).AddDays(-10)
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
If ($Folder.Value -like "%Systemdrive%*")
{
    Get-ChildItem C:\Inetpub\Logs\LogFiles -Recurse | Where-Object {$_.LastWriteTime -LT $Days -and $_.Name -like "*.LOG"} | Remove-Item -Force -ErrorAction SilentlyContinue
}
 Else
{   
    Get-ChildItem $Folder.Value -Recurse | Where-Object {$_.LastWriteTime -LT $Days -and $_.Name -like "*.LOG"} | Remove-Item -Force -ErrorAction SilentlyContinue
}

Get-ChildItem "$($env:ExchangeInstallPath)Logging" -Recurse | Where-Object {$_.LastWriteTime -LT $Days -and $_.Name -like "*.LOG"} | Remove-Item -Force -ErrorAction SilentlyContinue
