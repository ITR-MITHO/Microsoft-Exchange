<#

This script will start by exporting a list of all running Exchange services to a .csv-file that is placed on your desktop.
When you run the script, you will be prompted to Start or Stop Exchange services. 

The start function of this script will only be starting services found in the .csv-file, and set them to automatic startup. 
This way we prevent IMAP and POP3 from being started if they aren't needed


#>


# Check if PowerShell is started as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
Write-Host "Start PowerShell as an Administrator" -ForeGroundColor Yellow
Break
}


# Check if a list of running services exists, if not, then one will be created.
Add-PSSnapin *EXC*
$Path = "$Home\Desktop\ExchangeServices.csv"
$File = Test-Path $Path
If (-Not $File)
{
Get-Service | Where {$_.DisplayName -like "Microsoft Exchange*" -and $_.Status -EQ "Running"} | Select Name | Export-csv $Path -NoTypeInformation -Encoding Unicode
}

Function SERVICES { 
    $Function = Read-Host "
Please enter one of the below numbers to proceed:

    1 - STOP AND DISABLE RUNNING EXCHANGE SERVICES
    2 - START SERVICES THAT WAS DISABLED
    "
 If ($Function -EQ "1")
    {

Write-Host "

IMPORTANT - PLEASE READ:
If you choose to continue, all running Microsoft Exchange services on $env:computername will be stopped and disabled

" -ForegroundColor Yellow
$Confirm = Read-Host "Are you sure you want to continue? (Y/N)"
If ($Confirm -eq "Y")

{
$Services = Import-csv $Path
Foreach ($S in $Services)
{
Stop-Service -Name $S.Name -Force
Set-Service -Name $S.Name -StartupType Disabled
}
Write-Host "Services stopped and disabled.
Do NOT delete $Path if you want to start the correct services again with this script!" -ForegroundColor Yellow

    }
         }

If ($Function -EQ "2")
    {

$Services = Import-csv $Path
Foreach ($S in $Services)
{
Set-Service -Name $S.Name -StartupType Automatic
Start-Service -Name $S.Name
}
Write-Host "All Services found in $Path have been started again and set to automatic startup" -ForegroundColor Green

    }
        }

# This needs to be here to start the loop.
Services
