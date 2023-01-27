<#
IMPORTANT: Must be run in elevated Powershell
Will delete Exchange IIS-logs older then 14-days

#>

Import-Module WebAdministration
$Date = (Get-Date).AddDays(-14)
$Folder = Get-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory | Select Value
If ($Folder.Value -like "%Systemdrive%*")

{

    Get-ChildItem C:\Inetpub\Logs\LogFiles -Recurse | Where-Object {$_.LastWriteTime -lt $Date} | Remove-Item -Force
    

}

 Else

{
    
    Get-ChildItem $Folder.Value -Recurse | Where-Object {$_.LastWriteTime -lt $Date} | Remove-Item -Force

}
