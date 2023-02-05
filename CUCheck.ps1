<#

.DESCRIPTION
The script will check if a installed version of Exchange is supported or not.

.SYNOPSIS
By looking at the registry-database we can determine what version of Exchange is installed on the server.

- Exchange 2019 CU11+12      = SUPPORTED - EOL 14 October 2025
- Exchange 2016 CU23         = SUPPORTED - EOL 14 October 2025
- Exchange 2013 CU23         = SUPPORTED - EOL 11 April 2023
- Exchange 2010              = NOT SUPPORTED - EOL 13 October 2020
- Exchange 2007              = NOT SUPPORTED - EOL 11 April 2017

#>
Add-PSSnapin *EXC*
$Display = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "Microsoft Exchange Server 20*"}
$Exchange2019 = "Microsoft Exchange Server 2019 Cumulative Update 11"
$Exchange2016 = "Microsoft Exchange Server 2016 Cumulative Update 23"
$Exchange2013 = "Microsoft Exchange Server 2013 Cumulative Update 23"
$Exchange2010 = "Microsoft Exchange Server 2010"
$Exchange2007 = "Microsoft Exchange Server 2007"
$Servers = Get-ExchangeServer -Identity $ENV:COMPUTERNAME
$Output = $Display.DisplayName

Foreach ($Server in $Servers)
{
switch ($Output) {
    $Exchange2019 {
        Write-Host "$Output - SUPPORTED" -ForegroundColor Green
        break
    }
    $Exchange2016 {
        Write-Host "$Output - SUPPORTED" -ForegroundColor Green
        break
    }
    $Exchange2013 {
        Write-Host "$Output - SUPPORTED" -ForegroundColor Green
        break
    }
    $Exchange2010 {
        Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red
        break
    }
    $Exchange2007 {
        Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red
        break
}
    }
        }
