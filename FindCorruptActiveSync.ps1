$Computers = Get-ExchangeServer | Select Name
$properties = @(
    @{n='Error';e={$_.Properties[1].Value}}
)
$Date = (Get-Date).AddHours(-15)

Foreach ($Computer in $Computers)
{
Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=$Date; Id='1008'; ProviderName='MSExchange ActiveSync'}  | Select $properties |
Export-csv C:\Users\$ENV:Username\Desktop\Export1.csv -Append -NoTypeInformation
}

$csv = Import-Csv "C:\Users\$env:username\Desktop\Export1.csv"
Write-Host $csv
Foreach ($cell in $csv)
    {
    $cell -match '(Guid:\s[\S]{8}-[\S]{4}-[\S]{4}-[\S]{4}-[\S]{12})'
    Echo $Matches.Item(0) >> "C:\Users\$env:username\Desktop\GUIDS.csv"
    }
