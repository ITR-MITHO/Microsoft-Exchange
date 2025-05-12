$Path = Test-Path "C:\ITM8 - Scripts"
If (-Not $Path)
{
mkdir "C:\ITM8 - Scripts" | Out-Null
}

$PS = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/ITR-MITHO/Test-Scripts/refs/heads/main/PowerShell"
$Bat = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/ITR-MITHO/Test-Scripts/refs/heads/main/Bat"

$PS | Out-File -FilePath "C:\ITM8 - Scripts\ExchangePowerShell.ps1" -Encoding UTF8
$BAT | Out-File -FilePath "C:\ITM8 - Scripts\ExchangePowerShell.BAT" -Encoding ASCII
