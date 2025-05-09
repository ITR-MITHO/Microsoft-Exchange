<#

Creates a scheduled task, where only the .ps1 file and the user the task runs with should be changed.

#>

# Folder Paths
$folderPath = "C:\ITM8 - Scripts"
$batFilePath = Join-Path $folderPath "ExchangePowerShell.bat"
$ps1FilePath = Join-Path $folderPath "ExchangePowerShell.ps1"

# Create the folder if it doesn't exist
if (-not (Test-Path -Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
}

# Create the .bat file
$batContent = @'
CD "C:\ITM8 - Scripts"
PowerShell.exe -ExecutionPolicy Bypass -File .\ExchangePowerShell.ps1
'@
Set-Content -Path $batFilePath -Value $batContent -Encoding ASCII

# Create the .ps1 file 
$ps1Content = @'
Add-PSSnapin *EXC*
Import-Module ActiveDirectory
'@
Set-Content -Path $ps1FilePath -Value $ps1Content -Encoding UTF8

# Scheduled task params
$Name = Read-host "Enter Scheduled task name"
$batFilePath = "C:\ITM8 - Scripts\ExchangePowerShell.bat"

# Create the scheduled task
schtasks.exe /Create /TN "$Name" /TR "`"$batFilePath`"" /SC DAILY /ST 05:00 /RU "SYSTEM" /RL HIGHEST /F
