<#
.SYNOPSIS
    Exchange Server Daily Health Check

.DESCRIPTION  
    Version 1.1
    Performs a local health check on an Exchange Server.

    Checks:
    * C: drive free space
    * Exchange server component states
    * Message queue length
    * Backpressure events (last 24h)
    * DAG replication status
    * Exchange services
    * MAPI connectivity
    * Outlook MAPI/HTTP probe

.NOTES
    * Run in an elevated Exchange Management Shell
    * Run on each Exchange server individually
#>

# Ensure the script runs as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: Script must be run as Administrator." -ForegroundColor Red
    return
}

# Load required modules
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Add-PSSnapin *Exchange* -ErrorAction Stop
} catch {
    Write-Error "Failed to load required modules or snap-ins: $_"
    return
}

# --- Check C: drive free space ---
$spacePercent = Get-PSDrive C | ForEach-Object { [math]::Round($_.Free / ($_.Used + $_.Free) * 100, 2) }
Write-Host "`n=== Disk Space Check ===" -ForegroundColor Cyan
if ($spacePercent -lt 20) {
    Write-Host "C: drive has only $spacePercent% free space! Minimum required is 20%." -ForegroundColor Red
    Write-Host "Consider running IIS log cleanup: https://github.com/ITR-MITHO/Microsoft-Exchange/blob/main/EXCH_IISLogCleanup.ps1" -ForegroundColor Yellow
} else {
    Write-Host "C: drive space: OK ($spacePercent%)" -ForegroundColor Green
}

# --- Check message queue ---
Write-Host "`n=== Message Queue ===" -ForegroundColor Cyan
try {
    $queueCount = (Get-ExchangeServer | Get-Message -ErrorAction Stop).Count
    if ($queueCount -gt 100) {
        Write-Host "Warning: $queueCount messages in queue." -ForegroundColor Red
    } else {
        Write-Host "Message queue count: OK ($queueCount)" -ForegroundColor Green
    }
} catch {
    Write-Warning "Unable to get message queue count. $_"
}

# --- Check component states ---
Write-Host "`n=== Component States ===" -ForegroundColor Cyan
$components = Get-ServerComponentState -Identity $env:COMPUTERNAME |
    Where-Object { $_.Component -notin @("ForwardSyncDaemon", "ProvisioningRps") }

$inactive = $components | Where-Object { $_.State -eq "Inactive" }
if ($inactive) {
    Write-Host "Inactive Exchange components detected:" -ForegroundColor Red
    $inactive | Format-Table -AutoSize
} else {
    Write-Host "All Exchange components active." -ForegroundColor Green
}

# --- Check Exchange services ---
Write-Host "`n=== Exchange Services ===" -ForegroundColor Cyan
$serviceHealth = Test-ServiceHealth $env:COMPUTERNAME
$missing = $serviceHealth | Where-Object { $_.RequiredServicesRunning -ne $true }
if ($missing) {
    Write-Host "Some required Exchange services are not running:" -ForegroundColor Red
    $missing | Select-Object Role, ServicesNotRunning | Format-Table -AutoSize
} else {
    Write-Host "All required Exchange services are running." -ForegroundColor Green
}

# --- Check for backpressure events ---
Write-Host "`n=== Backpressure Events (last 24h) ===" -ForegroundColor Cyan
$backpressureIDs = @(15004, 15005, 15006, 15007)
$events = Get-WinEvent -FilterHashtable @{
    LogName = 'Application'
    ProviderName = 'MSExchangeTransport'
    ID = $backpressureIDs
    StartTime = (Get-Date).AddHours(-24)
} -ErrorAction SilentlyContinue

if ($events.Count -eq 0) {
    Write-Host "No backpressure events in the last 24 hours." -ForegroundColor Green
} else {
    Write-Host "Backpressure events found:" -ForegroundColor Red
    $events | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize
}

# --- MAPI connectivity ---
Write-Host "`n=== MAPI Connectivity ===" -ForegroundColor Cyan
$mapiResults = Test-MAPIConnectivity
$failedMapi = $mapiResults | Where-Object { $_.Result -eq "Failed" }
if ($failedMapi) {
    Write-Host "MAPI connectivity test failed:" -ForegroundColor Red
    $failedMapi | Format-Table -AutoSize
} else {
    Write-Host "MAPI connectivity: OK" -ForegroundColor Green
}

# --- Outlook connectivity ---
Write-Host "`n=== Outlook Connectivity ===" -ForegroundColor Cyan
$outlookResult = Test-OutlookConnectivity -ProbeIdentity OutlookMapiHttp.Protocol\OutlookMapiHttpSelfTestProbe -ErrorAction SilentlyContinue
if ($outlookResult -and $outlookResult.Result -eq "Failed") {
    Write-Host "Outlook MAPI/HTTP probe failed:" -ForegroundColor Red
    $outlookResult | Format-Table -AutoSize
} else {
    Write-Host "Outlook connectivity: OK" -ForegroundColor Green
}

# --- DAG replication ---
Write-Host "`n=== DAG Replication Health ===" -ForegroundColor Cyan
$dag = Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue
if ($dag) {
    $dagResults = Test-ReplicationHealth -Identity $env:COMPUTERNAME | Where-Object { $_.Result -like "*Failed*" }
    if ($dagResults) {
        Write-Host "DAG replication health issues found:" -ForegroundColor Red
        $dagResults | Select-Object Server, Check, Result | Format-Table -AutoSize
    } else {
        Write-Host "DAG replication: OK" -ForegroundColor Green
    }
} else {
    Write-Host "No DAG found on this server. Skipping DAG replication check." -ForegroundColor Yellow
}

# --- Wrap up ---
Write-Host "`nHealth check completed for $env:COMPUTERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
