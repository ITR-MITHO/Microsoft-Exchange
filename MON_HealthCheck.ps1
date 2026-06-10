<#
.SYNOPSIS
    Exchange Server Daily Health Check (Email Edition)
.DESCRIPTION
    Performs an local health check on an individual Exchange Server
    and outputs an HTML report into an e-mail
.NOTES
    Must be executed from an elevated Exchange Management Shell session.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *Exchange* -ErrorAction SilentlyContinue
}

$TargetServer = (Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop).Name
$MinPercent   = 20 # Available free space threshold
$ReportRows   = [System.Collections.Generic.List[string]]::new()

# Helper function to append rows to the HTML report table
function Add-ReportRow {
    param (
        [string]$Section,
        [string]$Check,
        [string]$Status, # 'OK', 'Warning', 'Critical'
        [string]$Details
    )
    $Class = switch ($Status) {
        'OK'       { 'status-ok' }
        'Warning'  { 'status-warn' }
        'Critical' { 'status-crit' }
        default    { '' }
    }
    $Row = "<tr class='$Class'><td>$Section</td><td>$Check</td><td>$Status</td><td>$Details</td></tr>"
    $ReportRows.Add($Row)
}

function Get-FreeSpacePercent {
    param ([string]$DriveLetter)
    $Drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$DriveLetter'" -ErrorAction SilentlyContinue
    if ($null -eq $Drive) { return $null }
    return [math]::Round(($Drive.FreeSpace / $Drive.Size) * 100, 2)
}

# ==========================================
# 1. DISK SPACE CHECK
# ==========================================
$CPercent = Get-FreeSpacePercent -DriveLetter "C:"
if ($null -ne $CPercent) {
    if ($CPercent -lt $MinPercent) {
        Add-ReportRow -Section "Disk" -Check "C: Drive Space" -Status "Critical" -Details "Only $CPercent% free space! Min threshold is $MinPercent%."
    } else {
        Add-ReportRow -Section "Disk" -Check "C: Drive Space" -Status "OK" -Details "Free space: $CPercent%"
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
            Add-ReportRow -Section "Disk" -Check "$Drive Drive Space" -Status "Critical" -Details "Only $Percent% free space! Threshold is $MinPercent%."
        } else {
            Add-ReportRow -Section "Disk" -Check "$Drive Drive Space" -Status "OK" -Details "Free space: $Percent%"
        }
    }
}

# ==========================================
# 2. SYSTEM RESOURCES (CPU & RAM)
# ==========================================
$Cpu = Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average
$CpuPercent = [math]::Round($Cpu.Average, 2)
if ($CpuPercent -gt 85) {
    Add-ReportRow -Section "Resources" -Check "CPU Usage" -Status "Warning" -Details "High CPU Utilization: $CpuPercent%"
} else {
    Add-ReportRow -Section "Resources" -Check "CPU Usage" -Status "OK" -Details "CPU usage is at $CpuPercent%"
}

$Os = Get-CimInstance Win32_OperatingSystem
$TotalRamGb = [math]::Round($Os.TotalVisibleMemorySize / 1MB, 2)
$FreeRamGb  = [math]::Round($Os.FreePhysicalMemory / 1MB, 2)
$UsedRamGb  = $TotalRamGb - $FreeRamGb
$RamPercent = [math]::Round(($UsedRamGb / $TotalRamGb) * 100, 2)
if ($RamPercent -gt 85) {
    Add-ReportRow -Section "Resources" -Check "RAM Usage" -Status "Warning" -Details "High RAM Utilization: $RamPercent% ($UsedRamGb GB used of $TotalRamGb GB)"
} else {
    Add-ReportRow -Section "Resources" -Check "RAM Usage" -Status "OK" -Details "RAM usage is at $RamPercent% ($UsedRamGb GB used of $TotalRamGb GB)"
}

# ==========================================
# 3. MESSAGE QUEUE
# ==========================================
try {
    $QueueCount = (Get-Queue -Server $TargetServer -ErrorAction Stop | Measure-Object -Property MessageCount -Sum).Sum
    if ($QueueCount -gt 100) {
        Add-ReportRow -Section "Queues" -Check "Message Queue Count" -Status "Warning" -Details "High queue count detected: $QueueCount messages."
    } else {
        Add-ReportRow -Section "Queues" -Check "Message Queue Count" -Status "OK" -Details "Queue count is normal: $QueueCount"
    }
} catch {
    Add-ReportRow -Section "Queues" -Check "Message Queue Count" -Status "Critical" -Details "Failed to query queues: $_"
}

# ==========================================
# 4. COMPONENT STATES
# ==========================================
$Components = Get-ServerComponentState -Identity $TargetServer |
    Where-Object { $_.Component -notin @("ForwardSyncDaemon", "ProvisioningRps") }

$Inactive = $Components | Where-Object { $_.State -eq "Inactive" }
if ($Inactive) {
    $CompNames = ($Inactive | ForEach-Object { "$($_.Component) ($($_.Requester))" }) -join ", "
    Add-ReportRow -Section "Components" -Check "Server Component States" -Status "Critical" -Details "Inactive components found: $CompNames"
} else {
    Add-ReportRow -Section "Components" -Check "Server Component States" -Status "OK" -Details "All relevant components are Active."
}

# ==========================================
# 5. EXCHANGE SERVICES
# ==========================================
$ServiceHealth = Test-ServiceHealth -Server $TargetServer -ErrorAction SilentlyContinue
$Missing = $ServiceHealth | Where-Object { $_.RequiredServicesRunning -ne $true }
if ($Missing) {
    $RolesFailed = ($Missing | ForEach-Object { $_.Role }) -join ", "
    Add-ReportRow -Section "Services" -Check "Required Services" -Status "Critical" -Details "Required services not running for roles: $RolesFailed"
} else {
    Add-ReportRow -Section "Services" -Check "Required Services" -Status "OK" -Details "All required services are running."
}

# ==========================================
# 6. BACKPRESSURE EVENTS (LAST 24H)
# ==========================================
$Events = @(Get-WinEvent -FilterHashtable @{
    LogName      = 'Application'
    ProviderName = 'MSExchangeTransport'
    ID           = @(15004, 15005, 15006, 15007)
    StartTime    = (Get-Date).AddHours(-24)
} -ErrorAction SilentlyContinue)

if ($Events.Count -eq 0) {
    Add-ReportRow -Section "Transport" -Check "Backpressure Events" -Status "OK" -Details "No backpressure events in the last 24 hours."
} else {
    Add-ReportRow -Section "Transport" -Check "Backpressure Events" -Status "Warning" -Details "$($Events.Count) backpressure events logged in last 24h. Inspect Application log (Event IDs 15004-15007)."
}

# ==========================================
# 7. DATABASE STATUS
# ==========================================
try {
    $Databases = Get-MailboxDatabase -Server $TargetServer -Status -ErrorAction Stop
    foreach ($Db in $Databases) {
        $IsDagDb = $Db.MasterServerOrAvailabilityGroup.Name -ne $TargetServer

        if (-not $IsDagDb) {
            if (-not $Db.Mounted) {
                Add-ReportRow -Section "Databases" -Check "DB: $($Db.Name)" -Status "Critical" -Details "Standalone database is UNMOUNTED!"
            } else {
                Add-ReportRow -Section "Databases" -Check "DB: $($Db.Name)" -Status "OK" -Details "Mounted (Standalone)"
            }
        } else {
            $MdbCopyStatus = Get-MailboxDatabaseCopyStatus -Identity "$($Db.Name)\$TargetServer" -ErrorAction SilentlyContinue
            if ($MdbCopyStatus.Status -eq "Mounted") {
                Add-ReportRow -Section "Databases" -Check "DB: $($Db.Name)" -Status "OK" -Details "Mounted (Active copy)"
            } elseif ($MdbCopyStatus.Status -eq "Healthy") {
                Add-ReportRow -Section "Databases" -Check "DB: $($Db.Name)" -Status "OK" -Details "Healthy (Passive copy)"
            } else {
                Add-ReportRow -Section "Databases" -Check "DB: $($Db.Name)" -Status "Critical" -Details "Copy status is [$($MdbCopyStatus.Status)] on this node!"
            }
        }
    }
} catch {
    Add-ReportRow -Section "Databases" -Check "Database Retrieval" -Status "Critical" -Details "Failed to retrieve status: $_"
}

# ==========================================
# 8. DAG REPLICATION
# ==========================================
$dag = Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue | Where-Object { $_.Servers -match $env:COMPUTERNAME }
if ($dag) {
    $dagResults = Test-ReplicationHealth -Identity $env:COMPUTERNAME | Where-Object { $_.Result -like "*Failed*" }
    if ($dagResults) {
        $FailedChecks = ($dagResults | ForEach-Object { $_.Check }) -join ", "
        Add-ReportRow -Section "Replication" -Check "DAG Health" -Status "Critical" -Details "Failed checks: $FailedChecks"
    } else {
        Add-ReportRow -Section "Replication" -Check "DAG Health" -Status "OK" -Details "DAG replication health validation passed."
    }
} else {
    Add-ReportRow -Section "Replication" -Check "DAG Health" -Status "OK" -Details "No DAG membership detected. Skipped."
}

# ==========================================
# HTML GENERATION & EMAIL TRANSMISSION
# ==========================================
$HtmlHead = @"
<style>
    body { font-family: Calibri, Arial, sans-serif; font-size: 14px; }
    table { border-collapse: collapse; width: 100%; max-width: 900px; margin-top: 15px; }
    th { background-color: #0078d4; color: white; text-align: left; padding: 10px; font-weight: bold; }
    td { padding: 8px; border: 1px solid #e0e0e0; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    .status-ok { color: #107c41; font-weight: bold; }
    .status-warn { background-color: #fff2cc; color: #b77c00; font-weight: bold; }
    .status-crit { background-color: #fde7e9; color: #a80000; font-weight: bold; }
    .footer { margin-top: 20px; font-size: 12px; color: #555555; }
</style>
"@

$HtmlBody = @"
<h2>Exchange Server Health Report: $TargetServer</h2>
<p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<table>
    <tr>
        <th>Section</th>
        <th>Check</th>
        <th>Status</th>
        <th>Details</th>
    </tr>
    $($ReportRows -join "`n")
</table>
<div class='footer'>
    <p>This is an automated health check execution.</p>
</div>
"@

# Construct finalized HTML payload
$FinalHtml = "<html><head>$HtmlHead</head><body>$HtmlBody</body></html>"

# --- Send Mail Configuration ---
$MailArgs = @{
    To         = "Exchange.Team@itm8.com"
    From       = "ITM8-EXCH@domain.com" # Change sender domain to match actual sender domain
    Subject    = "Daily Exchange Health Report - $TargetServer - $((Get-Date).ToShortDateString())"
    Body       = $FinalHtml
    BodyAsHtml = $true
    SmtpServer = "localhost" 
    Port     = 25
}

Send-MailMessage @MailArgs
