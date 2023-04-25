<# 

.DESCRIPTION  
Version 1.0
The script is designed to help you check how Exchange is feeling today.

* Can be run without editing
* Free space on C-drive
* Exchange ServerComponents
* MessageQueue higher than 100
* DAG Replication Test
* Microsoft Exchange Services
* MAPI Connectivity
* Outlook Connectivity

.NOTES
* Run in a elevated Exchange Shell
* Run on each server individually. The script doesn't check every Exchange Server there is automatically.

#>

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
write-host "Script is not running as Administrator" -ForegroundColor Yellow
Break
}


Import-Module ActiveDirectory
Add-PSSnapin *EXC*
# Free space on C:\
$Space = get-psdrive c | % { $_.free/($_.used + $_.free) } | % tostring p
If ($Space -lt "20 %")
{

Write-Host "There is $space left on the C-drive. It have needs atleast 20%
" -ForegroundColor Red

Write-host "
Run this script to clean up the drive.
https://raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/IISLogCleanup.ps1." -ForegroundColor Yellow

}
Else
{

Write-Host "Free space on C:\ is above 20%" -ForegroundColor Green

}


# MessageQueue
$Queue = (Get-ExchangeServer | Get-Message -ErrorAction SilentlyContinue).count
$Date = Get-Date -Format "dd-MM-yy HH:mm"
If ($Queue -GT 100)
{
Write-Host "Over 100 e-mails are in queue!" -ForegroundColor Red
}
Else
{
Write-Host "

Exchange Message queue is healthy.

" -ForegroundColor Green
}


# ComponentState
$Component = Get-ServerComponentState -Identity $env:computername | Where {$_.Component -NE "ForwardSyncDaemon" -and $_.Component -NE "ProvisioningRps"}
if ($Component | Where {$_.State -eq "inactive"})
{
Write-Host "Exchange componenets are inactive!" -ForegroundColor Red
$Component
}
Else
{
Write-Host "Exchange Server Components is running

" -ForegroundColor Green
}


# ServiceHealth
$ServiceHealth = Test-ServiceHealth $env:computername
if ($ServiceHealth | Where {$_.RequiredServicesRunning -NE $true})
{
Write-Host "Microsoft Exchange Services are not running!" -ForegroundColor Red
$ServiceHealth
}
else
{
Write-Host "Microsoft Exchange Services are running

" -ForegroundColor Green
}


# MapiConnectivity
$MAPIConnectivity = Test-MAPIConnectivity
If ($MAPIConnectivity | Where {$_.Result -EQ "Failed"})
{
Write-Host "MapiConnectivity failed." -ForegroundColor Red
$MapiConnectivity
}
Else
{
Write-Host "MAPIConnectivityTest passed

" -ForegroundColor Green
}


# OutlookConnectivity
$OutlookConnectivity = Test-OutlookConnectivity -ProbeIdentity OutlookMapiHttp.Protocol\OutlookMapiHttpSelfTestProbe
If ($MAPIConnectivity  -EQ "Failed")
{
Write-Host "OutlookConnectivity failed." -ForegroundColor Red
$MapiConnectivity
}
Else
{
Write-Host "OutlookConnectivity passed

" -ForegroundColor Green
}


# DAGReplicationHealth
$DAGTest = Test-ReplicationHealth $env:computername | Where {$_.Result -like "*failed*"} | Select Server, Check, Result
$DAG = Get-DatabaseAvailabilityGroup
If ($DAG)
{

Write-Host "Exchange DAG Found.. Testing replication" -ForegroundColor Yellow
If (-Not $DagTEST)
{
Write-Host "DAG Replication test passed" -ForegroundColor Green
}
{

}
}
Else
{
Write-Host "No Exchange DAG found, skipping replication check.
" -ForegroundColor Yellow
}

sleep 5

If ($DAG -ne $null)
{

if ($DAGTest -ne $null)
{
cls
    Write-Host "Exchange DAG replication is unhealthy! (Test-ReplicationHealth)
    
    " -ForegroundColor Red
}

}

Write-Host "Healthcheck completed." -ForegroundColor Yellow
