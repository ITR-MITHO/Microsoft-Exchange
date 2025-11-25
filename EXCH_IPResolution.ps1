<#

.SYNOPSIS
Resolves client IP addresses from Exchange message tracking logs.

.DESCRIPTION
Collects 'Receive' event logs from all Exchange servers within the last 90 days,
resolves each unique OriginalClientIP to a hostname, and exports the data to a CSV.
Unresolvable IPs are listed as 'Unresolved'.

.OUTPUTS
A .csv-file will be placed on your desktop named SenderResolution.csv

#>
Add-PSSnapin *EXC*
$Data = Get-ExchangeServer |
Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddDays(-90) -EventId Receive |
Select-Object Sender, OriginalClientIP, MessageSubject, Timestamp, ConnectorID

$Grouped = $Data | Group-Object OriginalClientIP
$Results = foreach ($Group in $Grouped) {

    $IP = $Group.Name
    $Count = $Group.Count

    # DNS resolution once per IP
    try {
        $ResolvedName = [System.Net.Dns]::GetHostEntry($IP).HostName
    } catch {
        $ResolvedName = "Unresolved"
    }

    $Entry = $Group.Group | Select-Object -First 1
    [PSCustomObject]@{
        TimeStamp        = $Entry.Timestamp
        OriginalClientIP = $IP
        IPCount          = $Count
        Hostname         = $ResolvedName
        Sender           = $Entry.Sender
        Connector        = $Entry.ConnectorID
        Subject          = $Entry.MessageSubject
    }
}
# Export
$Results | Export-Csv -Path "$home\Desktop\SenderResolution.csv" -NoTypeInformation -Encoding Unicode
Write-Host "Export can be found here: $home\Desktop\SenderResolution.csv" -ForegroundColor Green
