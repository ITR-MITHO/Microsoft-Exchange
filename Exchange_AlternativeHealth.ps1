<#
.SYNOPSIS
    Exchange Server Daily Health Check (Silent Edition)
.DESCRIPTION
    Performs an optimized local health check on an individual Exchange Server. 
    Only outputs data if an issue is detected
.NOTES
    Must be executed from an elevated Exchange Management Shell session.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *Exchange* -ErrorAction SilentlyContinue
}

$TargetServer = (Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop).Name
$MinPercent   = 20 # Available free space threshhold
$IssuesFound  = $false

function Get-FreeSpacePercent {
    param ([string]$DriveLetter)
    $Drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$DriveLetter'" -ErrorAction SilentlyContinue
    if ($null -eq $Drive) { return $null }
    return [math]::Round(($Drive.FreeSpace / $Drive.Size) * 100, 2)
}

# --- Disk Space Check ---
$CPercent = Get-FreeSpacePercent -DriveLetter "C:"
if ($null -ne $CPercent -and $CPercent -lt $MinPercent) {
    Write-Host "C: drive has only $CPercent% free space! Minimum required is $MinPercent%." -ForegroundColor Red
    Write-Host "Consider running IIS log cleanup: https://github.com/ITR-MITHO/Microsoft-Exchange/blob/main/EXCH_IISLogCleanup.ps1" -ForegroundColor Yellow
    $IssuesFound = $true
}

$DrivesToCheck = [System.Collections.Generic.HashSet[string]]::new()
Get-MailboxDatabase -Server $TargetServer -Status -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.EdbFilePath.PathName -match '^([A-Za-z]:)') { [void]$DrivesToCheck.Add($Matches[1]) }
    if ($_.LogFolderPath.PathName -match '^([A-Za-z]:)') { [void]$DrivesToCheck.Add($Matches[1]) }
}

foreach ($Drive in $DrivesToCheck) {
    if ($Drive -eq 'C:') { continue }
    $Percent = Get-FreeSpacePercent -DriveLetter $Drive
    if ($null -ne $Percent -and $Percent -lt $MinPercent) {
        Write-Host "$Drive drive has only $Percent% free space! Threshold is $MinPercent%." -ForegroundColor Red
        $IssuesFound = $true
    }
}

# --- System Resources (CPU & RAM) ---
$Cpu = Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average
$CpuPercent = [math]::Round($Cpu.Average, 2)

if ($CpuPercent -gt 85) {
    Write-Host "CPU Usage: $CpuPercent%" -ForegroundColor Red
    $IssuesFound = $true
}

$Os = Get-CimInstance Win32_OperatingSystem
$TotalRamGb = [math]::Round($Os.TotalVisibleMemorySize / 1MB, 2)
$FreeRamGb  = [math]::Round($Os.FreePhysicalMemory / 1MB, 2)
$UsedRamGb  = $TotalRamGb - $FreeRamGb
$RamPercent = [math]::Round(($UsedRamGb / $TotalRamGb) * 100, 2)

if ($RamPercent -gt 85) {
    Write-Host "RAM Usage: $RamPercent% ($UsedRamGb GB used of $TotalRamGb GB)" -ForegroundColor Red
    $IssuesFound = $true
}

# --- Check message queue ---
try {
    $QueueCount = (Get-Queue -Server $TargetServer -ErrorAction Stop | Measure-Object -Property MessageCount -Sum).Sum
    if ($QueueCount -gt 100) {
        Write-Host "Warning: $QueueCount messages in local queues." -ForegroundColor Red
        $IssuesFound = $true
    }
} catch {
    Write-Warning "Unable to calculate local message queue count: $_"
    $IssuesFound = $true
}

# --- Check component states ---
$Components = Get-ServerComponentState -Identity $TargetServer |
    Where-Object { $_.Component -notin @("ForwardSyncDaemon", "ProvisioningRps") }

$Inactive = $Components | Where-Object { $_.State -eq "Inactive" }
if ($Inactive) {
    Write-Host "Inactive Exchange components detected:" -ForegroundColor Red
    $Inactive | Select-Object Component, State, Requester | Format-Table -AutoSize
    $IssuesFound = $true
}

# --- Check Exchange services ---
$ServiceHealth = Test-ServiceHealth -Server $TargetServer -ErrorAction SilentlyContinue
$Missing = $ServiceHealth | Where-Object { $_.RequiredServicesRunning -ne $true }
if ($Missing) {
    Write-Host "Some required Exchange services are not running:" -ForegroundColor Red
    $Missing | Select-Object Role, RequiredServicesRunning | Format-Table -AutoSize
    $IssuesFound = $true
}

# --- Check for backpressure events ---
$Events = @(Get-WinEvent -FilterHashtable @{
    LogName      = 'Application'
    ProviderName = 'MSExchangeTransport'
    ID           = @(15004, 15005, 15006, 15007)
    StartTime    = (Get-Date).AddHours(-24)
} -ErrorAction SilentlyContinue)

if ($Events.Count -gt 0) {
    Write-Host "$($Events.Count) Backpressure events found:" -ForegroundColor Red
    $Events | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize
    $IssuesFound = $true
}

# --- Database Mount & Copy Status ---
try {
    $Databases = Get-MailboxDatabase -Server $TargetServer -Status -ErrorAction Stop
    foreach ($Db in $Databases) {
        $IsDagDb = $Db.MasterServerOrAvailabilityGroup.Name -ne $TargetServer

        if (-not $IsDagDb) {
            if (-not $Db.Mounted) {
                Write-Host "CRITICAL: Standalone database [$($Db.Name)] is UNMOUNTED!" -ForegroundColor Red
                $IssuesFound = $true
            }
        } else {
            $MdbCopyStatus = Get-MailboxDatabaseCopyStatus -Identity "$($Db.Name)\$TargetServer" -ErrorAction SilentlyContinue
            if ($MdbCopyStatus.Status -notin @("Mounted", "Healthy")) {
                Write-Host "CRITICAL: Database [$($Db.Name)] copy status is [$($MdbCopyStatus.Status)] on this server!" -ForegroundColor Red
                $IssuesFound = $true
            }
        }
    }
} catch {
    Write-Warning "Unable to retrieve database status: $_"
    $IssuesFound = $true
}

# --- DAG replication ---
$dag = Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue | Where-Object { $_.Servers -match $env:COMPUTERNAME }

if ($dag) {
    $dagResults = Test-ReplicationHealth -Identity $env:COMPUTERNAME | Where-Object { $_.Result -like "*Failed*" }
    if ($dagResults) {
        Write-Host "DAG replication health issues found:" -ForegroundColor Red
        $dagResults | Select-Object Server, Check, Result | Format-Table -AutoSize
        $IssuesFound = $true
    }
}

# --- Final Output ---
if (-not $IssuesFound) {
    Write-Host "All tests have passed, the server is: HEALTHY" -ForegroundColor Green
}
