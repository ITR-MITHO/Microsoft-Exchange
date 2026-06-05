<#
.SYNOPSIS
    Exchange Server Daily Health Check
.DESCRIPTION
    Performs an optimized local health check on an individual Exchange Server.
.NOTES
    Must be executed from an elevated Exchange Management Shell session.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *Exchange* -ErrorAction SilentlyContinue
}

$TargetServer = (Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop).Name
$MinPercent   = 20

function Get-FreeSpacePercent {
    param ([string]$DriveLetter)
    $Drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$DriveLetter'" -ErrorAction SilentlyContinue
    if ($null -eq $Drive) { return $null }
    return [math]::Round(($Drive.FreeSpace / $Drive.Size) * 100, 2)
}

# --- Disk Space Check ---
Write-Host "`n=== Disk Space Check ===" -ForegroundColor Cyan

# Check C: Drive
$CPercent = Get-FreeSpacePercent -DriveLetter "C:"
if ($null -ne $CPercent) {
    if ($CPercent -lt $MinPercent) {
        Write-Host "C: drive has only $CPercent% free space! Minimum required is $MinPercent%." -ForegroundColor Red
        Write-Host "Consider running IIS log cleanup: https://github.com/ITR-MITHO/Microsoft-Exchange/blob/main/EXCH_IISLogCleanup.ps1" -ForegroundColor Yellow
    } else {
        Write-Host "C: drive space: OK ($CPercent%)" -ForegroundColor Green
    }
}
$DrivesToCheck = [System.Collections.Generic.HashSet[string]]::new()
Get-MailboxDatabase -Server $TargetServer -Status -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.EdbFilePath.PathName -match '^([A-Za-z]:)') { [void]$DrivesToCheck.Add($Matches[1]) }
    if ($_.LogFolderPath.PathName -match '^([A-Za-z]:)') { [void]$DrivesToCheck.Add($Matches[1]) }
}

foreach ($Drive in $DrivesToCheck) {
    if ($Drive -eq 'C:') { continue }
    $Percent = Get-FreeSpacePercent -DriveLetter $Drive
    if ($null -ne $Percent) {
        if ($Percent -lt $MinPercent) {
            Write-Host "$Drive drive has only $Percent% free space! Threshold is $MinPercent%." -ForegroundColor Red
        } else {
            Write-Host "$Drive drive space: OK ($Percent%)" -ForegroundColor Green
        }
    }
}

# --- Check message queue ---
Write-Host "`n=== Message Queue ===" -ForegroundColor Cyan
try {
    # Isolate strictly to the local server node context
    $QueueCount = (Get-Queue -Server $TargetServer -ErrorAction Stop | Measure-Object -Property MessageCount -Sum).Sum
    if ($QueueCount -gt 100) {
        Write-Host "Warning: $QueueCount messages in local queues." -ForegroundColor Red
    } else {
        Write-Host "Local message queue count: OK ($QueueCount)" -ForegroundColor Green
    }
} catch {
    Write-Warning "Unable to calculate local message queue count: $_"
}

# --- Check component states ---
Write-Host "`n=== Component States ===" -ForegroundColor Cyan
$Components = Get-ServerComponentState -Identity $TargetServer |
    Where-Object { $_.Component -notin @("ForwardSyncDaemon", "ProvisioningRps") }

$Inactive = $Components | Where-Object { $_.State -eq "Inactive" }
if ($Inactive) {
    Write-Host "Inactive Exchange components detected:" -ForegroundColor Red
    $Inactive | Select-Object Component, State, Requester | Format-Table -AutoSize
} else {
    Write-Host "All Exchange components active." -ForegroundColor Green
}

# --- Check Exchange services ---
Write-Host "`n=== Exchange Services ===" -ForegroundColor Cyan
$ServiceHealth = Test-ServiceHealth -Server $TargetServer -ErrorAction SilentlyContinue
$Missing = $ServiceHealth | Where-Object { $_.RequiredServicesRunning -ne $true }
if ($Missing) {
    Write-Host "Some required Exchange services are not running:" -ForegroundColor Red
    $Missing | Select-Object Role, RequiredServicesRunning | Format-Table -AutoSize
} else {
    Write-Host "All required Exchange services are running." -ForegroundColor Green
}

# --- Check for backpressure events ---
Write-Host "`n=== Backpressure Events (last 24h) ===" -ForegroundColor Cyan
# Forcing array literal coercion @(...) to protect item counting mechanics
$Events = @(Get-WinEvent -FilterHashtable @{
    LogName      = 'Application'
    ProviderName = 'MSExchangeTransport'
    ID           = @(15004, 15005, 15006, 15007)
    StartTime    = (Get-Date).AddHours(-24)
} -ErrorAction SilentlyContinue)

if ($Events.Count -eq 0) {
    Write-Host "No backpressure events in the last 24 hours." -ForegroundColor Green
} else {
    Write-Host "$($Events.Count) Backpressure events found:" -ForegroundColor Red
    $Events | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize
}

# --- DAG replication ---
Write-Host "`n=== DAG Replication Health ===" -ForegroundColor Cyan

# Filter the DAG query to ensure the local server is actually a member
$dag = Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue | Where-Object { $_.Servers -match $env:COMPUTERNAME }

if ($dag) {
    $dagResults = Test-ReplicationHealth -Identity $env:COMPUTERNAME | Where-Object { $_.Result -like "*Failed*" }
    
    if ($dagResults) {
        Write-Host "DAG replication health issues found:" -ForegroundColor Red
        $dagResults | Select-Object Server, Check, Result | Format-Table -AutoSize
    } else {
        Write-Host "DAG replication: OK" -ForegroundColor Green
    }
} else {
    Write-Host "No DAG found containing this server. Skipping DAG replication check." -ForegroundColor Yellow
}
Write-Host "`nHealth check completed for $TargetServer at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
