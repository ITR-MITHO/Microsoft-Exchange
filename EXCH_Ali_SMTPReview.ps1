<#
    .SYNOPSIS
    SMTP-Review.ps1

    .DESCRIPTION
    Script is intended to help determine servers that are using an Exchange server to connect and send email.
    This is especially pertinent in a decommission scenario, where the logs are to be checked to ensure that
    all SMTP traffic has been moved to the correct endpoint.

    .LINK
    www.alitajran.com/find-ip-addresses-using-exchange-smtp-relay

    .NOTES
    Written by: ALI TAJRAN
    Website:    www.alitajran.com
    LinkedIn:   linkedin.com/in/alitajran

    .CHANGELOG
    V1.00, 04/05/2021 - Initial version
    V2.00, 03/28/2023 - Rewrite script to retrieve results faster
#>

# Clears the host console to make it easier to read output
Clear-Host

# Sets the path to the directory containing the log files to be processed
$logFilePath = "C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive\*.log"

# Sets the path to the output file that will contain the unique IP addresses
$Output = "C:\temp\IPAddresses.txt"

# Gets a list of the log files in the specified directory
$logFiles = Get-ChildItem $logFilePath

# Gets the number of log files to be processed
$count = $logFiles.Count

# Initializes an array to store the unique IP addresses
$ips = foreach ($log in $logFiles) {

    # Displays progress information
    $percentComplete = [int](($logFiles.IndexOf($log) + 1) / $count * 100)
    $status = "Processing $($log.FullName) - $percentComplete% complete ($($logFiles.IndexOf($log)+1) of $count)"
    Write-Progress -Activity "Collecting Log details" -Status $status -PercentComplete $percentComplete

    # Displays the name of the log file being processed
    Write-Host "Processing Log File $($log.FullName)" -ForegroundColor Magenta

    # Reads the content of the log file, skipping the first five lines
    $fileContent = Get-Content $log | Select-Object -Skip 5

    # Loops through each line in the log file
    foreach ($line in $fileContent) {

        # Extracts the IP address from the socket information in the log line
        $socket = $line.Split(',')[5]
        $ip = $socket.Split(':')[0]

        # Adds the IP address to the $ips array
        $ip
    }
}

# Processing progress
Write-Progress -Activity "Processing IP Addresses" -Status "Counting IP addresses"

# Rens liste
$cleanIps = $ips | ForEach-Object { $_.ToString().Trim() } | Where-Object { $_ -and $_ -ne '-' }

# Count IPs
$ipCounts = @{}
foreach ($ip in $cleanIps) {
    if ($ipCounts.ContainsKey($ip)) { $ipCounts[$ip]++ } else { $ipCounts[$ip] = 1 }
}

# Build results objects including reverse DNS lookup
$results = $ipCounts.GetEnumerator() | ForEach-Object {
    $hostname = (Resolve-DnsName $_.Name -ErrorAction SilentlyContinue |
                 Select-Object -ExpandProperty NameHost -First 1)

    [PSCustomObject]@{
        IP       = $_.Name
        Hostname = $hostname
        Count    = $_.Value
    }
}

# Sort by count (descending)
$resultsSorted = $results | Sort-Object -Property Count -Descending

# Output to console
Write-Host "SMTP Senders (IP / Hostname / Count):" -ForegroundColor Cyan
$resultsSorted | Format-Table -AutoSize

# Write plain output
$Output = "C:\temp\IPAddresses.txt"
$OutputWithCounts = "C:\temp\SMTP-IPs-with-counts.csv"

# Save table to CSV
$resultsSorted | Export-Csv -Path $OutputWithCounts -Encoding UTF8 -NoTypeInformation

# Save just IP list
$resultsSorted | Select-Object -ExpandProperty IP | Out-File $Output -Encoding UTF8

Write-Host "`nSaved sorted CSV to $OutputWithCounts" -ForegroundColor Green
Write-Host "Saved unique IP list to $Output" -ForegroundColor Green
