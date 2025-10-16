<#

The script is designed to resolve IP's that uses Exchange to relay e-mails. 
It will collect MessageTrackingLogs from the last 5 days and try to resolve the IPs. 

If it cannot resolve the IP, it will be listed as 'Unresolved' 

.OUTPUT
A .csv-file will be placed on your desktop named SenderResolution.csv

#>
$Results = @()
$Data = Get-ExchangeServer | Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddDays(-5) -EventId Receive | 
Select-Object Sender, @{Name='Recipients';Expression={$_.Recipients}}, OriginalClientIP, MessageSubject, Timestamp

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
        Recipients       = ($Entry.Recipients -join ", ")
        OriginalClientIP = $Entry.OriginalClientIP
        Hostname         = $ResolvedName
        Subject          = $Entry.MessageSubject
    }
}

# Export to CSV
$Results | Export-Csv -Path $home\Desktop\SenderResolution.csv -NoTypeInformation -Encoding Unicode
Write-Host "Export can be found here $home\Desktop\SenderResolution.csv" -ForegroundColor Green
