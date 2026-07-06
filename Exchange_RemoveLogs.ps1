<#

Deletes .LOG files older than 10 days in the following directories: 
inetpub\logs\LogFiles
$($env:ExchangeInstallPath)Logging\*

#>
Import-Module WebAdministration
$CutoffDate = (Get-Date).AddDays(-10)

# Change IIS log rollover to Hourly instead of daily
$RollOver = (Get-WebConfigurationProperty -Filter /system.applicationHost/sites/siteDefaults/logFile -Name Period).Value
If ($RollOver -eq "Daily") {
    Set-WebConfigurationProperty -Filter /system.applicationHost/sites/siteDefaults/logFile -Name "period" -Value "Hourly"
}

$Folder = (Get-ItemProperty "IIS:\Sites\Default Web Site" -Name logFile.directory).Value
$IisPath = if ($Folder -like "%Systemdrive%*") { "C:\Inetpub\Logs\LogFiles" } else { $Folder }

Get-ChildItem -Path $IisPath -Filter "*.log" -Recurse -File | 
    Where-Object { $_.LastWriteTime -lt $CutoffDate } | 
    Remove-Item -Force -ErrorAction SilentlyContinue

Get-ChildItem -Path "$env:ExchangeInstallPath\Logging" -Filter "*.log" -Recurse -File | 
    Where-Object { $_.LastWriteTime -lt $CutoffDate } | 
    Remove-Item -Force -ErrorAction SilentlyContinue
