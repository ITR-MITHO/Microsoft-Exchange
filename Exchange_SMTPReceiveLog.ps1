<#
.SYNOPSIS
    Resolves client IP addresses from Exchange SMTP Receive Protocol Logs.
.DESCRIPTION
    Parses raw SMTP Protocol log files using Select-String
    Extracts IP, Sender, and Connector.
#>

$DaysBack = 14
$StartDate = (Get-Date).AddDays(-$DaysBack)
$ExportPath = Join-Path $home "Desktop\ProtocolLogIPs.csv"

# Locate default Exchange protocol log paths
$ExchangePath = $env:ExchangeInstallPath
if (-not $ExchangePath) {
    Write-Warning "ExchangeInstallPath environment variable not found. Run this on an Exchange server."
    exit
}

$LogPaths = @(
    Join-Path $ExchangePath "TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive"
    Join-Path $ExchangePath "TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive"
)

Write-Host "Locating protocol logs..." -ForegroundColor Cyan
$LogFiles = Get-ChildItem -Path $LogPaths -Filter "*.log" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $StartDate }

if (-not $LogFiles) {
    Write-Warning "No log files found in the last $DaysBack days."
    exit
}

$UniqueIPs = @{}
$ResultsList = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Parsing $($LogFiles.Count) log files using optimized string matching..." -ForegroundColor Yellow

# Select-String is vastly faster for large text files.
# We target 'MAIL FROM:' to capture the sender and the remote IP in a single pass.
$Matches = $LogFiles | Select-String -Pattern "MAIL FROM:" -SimpleMatch

foreach ($Match in $Matches) {
    # Expected Exchange Protocol Log format:
    # date-time,connector-id,session-id,sequence-number,local-endpoint,remote-endpoint,event,data,context
    $LineParts = $Match.Line -split ','

    if ($LineParts.Count -ge 8) {
        $RawIP = $LineParts[5]
        
        # Strip the dynamic port number (e.g., 192.168.1.10:64856 -> 192.168.1.10)
        # Handles IPv4 and IPv6 formats
        $IP = if ($RawIP -match "\[?(.*?)\]?:\d+$") { $matches[1] } else { $RawIP }

        if ([string]::IsNullOrWhiteSpace($IP)) { continue }

        # Extract Sender from the data column (e.g., MAIL FROM:<sender@domain.com>)
        $RawSender = $LineParts[7]
        $Sender = if ($RawSender -match "<(.*?)>") { $matches[1] } else { $RawSender -replace "MAIL FROM:","" }

        $UniqueIPs[$IP]++

        # Storing metadata. Note: Subject is not natively visible in this stage of the protocol log.
        $UniqueIPs["$IP-Metadata"] = [PSCustomObject]@{
            TimeStamp        = [datetime]::Parse($LineParts[0]).ToString("yyyy-MM-dd HH:mm:ss")
            OriginalClientIP = $IP
            Sender           = $Sender
            Connector        = $LineParts[1]
            Subject          = "N/A (Protocol Logs)"
        }
    }
}

$UniqueIPCount = ($UniqueIPs.Keys | Where-Object { $_ -notmatch '-Metadata$' }).Count
Write-Host "`nProcessing DNS resolution for $UniqueIPCount unique IPs..." -ForegroundColor Cyan

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
    $ResultsList | Select-Object TimeStamp, OriginalClientIP, IPCount, Hostname, Sender, Connector |
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding Unicode

    Write-Host "`nAnalysis complete! Results exported to: $ExportPath" -ForegroundColor Green
} else {
    Write-Host "No matching data found." -ForegroundColor Yellow
}
