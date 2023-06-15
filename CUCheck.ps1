<#

.DESCRIPTION
The script will check if a installed version of Exchange is supported or not.

.SYNOPSIS
By looking at the registry-database we can determine what version of Exchange is installed on the server.

- Exchange 2019 CU12+13      = SUPPORTED - EOL 14 October 2025
- Exchange 2016 CU23         = SUPPORTED - EOL 14 October 2025
- Exchange 2013 CU23         = NOT SUPPORTED - EOL 11 April 2023
- Exchange 2010              = NOT SUPPORTED - EOL 13 October 2020
- Exchange 2007              = NOT SUPPORTED - EOL 11 April 2017

#>

<# Arrays #>
Add-PSSnapin *EXC*
$Display = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "Microsoft Exchange Server 20*"}
$Exchange2019 = "Microsoft Exchange Server 2019 Cumulative Update 12"
$Exchange2016 = "Microsoft Exchange Server 2016 Cumulative Update 23"
$Exchange2013 = "Microsoft Exchange Server 2013"
$Exchange2010 = "Microsoft Exchange Server 2010"
$Exchange2007 = "Microsoft Exchange Server 2007"
$Servers = Get-ExchangeServer -Identity $Env:ComputerName
$Output = $Display.DisplayName

<# Exchange Server 2019 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*Microsoft Exchange Server 2019*")
{

if ($Display.DisplayName -GE "$Exchange2019")
{

Write-Host "$Output - SUPPORTED" -ForegroundColor Green

}
Else
{

Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}

   }
        }
    

<# Exchange Server 2016 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*Microsoft Exchange Server 2016*")
{

if ($Display.DisplayName -GE "$Exchange2016")
{

Write-Host "$Output - SUPPORTED" -ForegroundColor Green

}
Else
{

Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}

   }
        }


<# Exchange Server 2013 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*Microsoft Exchange Server 2013*")
{

if ($Display.DisplayName -Like "*$Exchange2013*")
{

        Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}

   }

<# Exchange Server 2010 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*$Exchange2010*")

{

        Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}
    
        }


<# Exchange Server 2007 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*$Exchange2007*")

{

        Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}
    
        }
