<#

.SYNOPSIS
Resolves client IP addresses from Exchange message tracking logs.

.DESCRIPTION
Collects 'Receive' event logs from all Exchange servers within the last 5 days,
resolves each unique OriginalClientIP to a hostname, and exports the data to a CSV.
Unresolvable IPs are listed as 'Unresolved'.

.OUTPUTS
A .csv-file will be placed on your desktop named SenderResolution.csv

#>
Add-PSSnapin *EXC*
$Results = @()
$Data = Get-ExchangeServer | Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddDays(-5) -EventId Receive | 
Select-Object Sender, OriginalClientIP, MessageSubject, Timestamp

foreach ($Entry in $Data) {
$ResolvedName = $null
try {
    $ResolvedName = [System.Net.Dns]::GetHostEntry($Entry.OriginalClientIP).HostName
    }
catch {
    $ResolvedName = "Unresolved"
    }
    $Results += [PSCustomObject]@{
        TimeStamp        = $Entry.TimeStamp
        Sender           = $Entry.Sender
        OriginalClientIP = $Entry.OriginalClientIP
        Hostname         = $ResolvedName
        Subject          = $Entry.MessageSubject
    }
}

# Export to CSV
$Results | Export-Csv -Path $home\Desktop\SenderResolution.csv -NoTypeInformation -Encoding Unicode
Write-Host "Export can be found here $home\Desktop\SenderResolution.csv" -ForegroundColor Green
