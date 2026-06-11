<#
.SYNOPSIS
    Resolves client IP addresses from Exchange message tracking logs (90-Day Scale Optimized).
.DESCRIPTION
    Collects 'Receive' event logs from all transport nodes within the last 90 days,
    deduplicates client IPs on-the-fly, resolves hostnames, and exports the data.
    Standardizes timestamps to prevent Excel formatting corruption.
.OUTPUTS
    $home\Desktop\MessageTraceIPs.csv
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

$StartDate = (Get-Date).AddDays(-90)
$ExportPath = Join-Path $home "Desktop\MessageTraceIPs.csv"

Write-Host "Gathering transport servers..." -ForegroundColor Cyan
$Servers = Get-ExchangeServer | Where-Object { $_.IsHubTransportServer -or $_.IsMailboxServer }


$UniqueIPs = @{}
$ResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Parsing Message Tracking Logs..." -ForegroundColor Yellow

foreach ($Server in $Servers) {
    Write-Host "Querying node: $($Server.Name)..." -ForegroundColor DarkCyan
    try {
        # Streaming logs server-by-server rather than choking a giant multi-server pipeline
        $Logs = Get-MessageTrackingLog -Server $Server.Name -ResultSize Unlimited -Start $StartDate -EventId Receive -ErrorAction Stop
        
        foreach ($Log in $Logs) {
            # Trim potential trailing whitespace/tabs from the tracking logs right at ingestion
            $IP = if ($Log.OriginalClientIP) { $Log.OriginalClientIP.ToString().Trim() } else { $null }
            if ([string]::IsNullOrEmpty($IP)) { continue }

            $UniqueIPs[$IP]++
            $UniqueIPs["$IP-Metadata"] = [PSCustomObject]@{
                # FIX: Force a clean, universal string date format that Excel cannot break or misinterpret
                TimeStamp        = $Log.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
                OriginalClientIP = $IP
                Sender           = $Log.Sender
                Connector        = $Log.ConnectorID
                Subject          = $Log.MessageSubject
            }
        }
    } catch {
        Write-Warning "Skipped processing for server $($Server.Name). Reason: $_"
    }
}

$UniqueIPCount = ($UniqueIPs.Keys | Where-Object { $_ -notmatch '-Metadata$' }).Count
Write-Host "`nProcessing DNS resolution for $UniqueIPCount unique client IPs..." -ForegroundColor Cyan

$DnsCache = @{}
foreach ($Key in $UniqueIPs.Keys) {
    if ($Key -match '-Metadata$') { continue }
    $IP = $Key
    $Metadata = $UniqueIPs["$IP-Metadata"]
    if (-not $DnsCache.ContainsKey($IP)) {
        try {
            $DnsCache[$IP] = [System.Net.Dns]::GetHostEntry($IP).HostName
        } catch {
            $DnsCache[$IP] = "Unresolved"
        }
    }

    $Metadata | Add-Member -MemberType NoteProperty -Name "IPCount" -Value $UniqueIPs[$IP] -Force
    $Metadata | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $DnsCache[$IP] -Force

    $ResultsList.Add($Metadata)
}

if ($ResultsList.Count -gt 0) {
    $ResultsList | Select-Object TimeStamp, OriginalClientIP, IPCount, Hostname, Sender, Connector, Subject | 
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding Unicode
    
    Write-Host "`nAnalysis complete! Results exported cleanly to: $ExportPath" -ForegroundColor Green
} else {
    Write-Host "No message tracking rows containing valid client IPs discovered." -ForegroundColor Yellow
}
