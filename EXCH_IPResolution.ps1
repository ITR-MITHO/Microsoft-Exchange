<#
The script is designed to resolve IP's that uses Exchange to relay e-mails. 
It will collect MessageTrackingLogs from the last 30 days and try to resolve the IPs. 

If it cannot resolve the IP, it will be listed as 'Unresolved' 

.OUTPUT
A .csv-file will be placed on your desktop named ClientIP_Resolution.csv


#>

$Results = @()

$Data = Get-ExchangeServer | Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddDays(-30) -EventId Receive | 
    Select-Object Sender, @{Name='Recipients';Expression={$_.Recipients}}, OriginalClientIP, MessageSubject

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
        Subject          = $Entry.MessageSubject
        OriginalClientIP = $Entry.OriginalClientIP
        Hostname         = $ResolvedName
    }
}

# Export to CSV
$Results | Export-Csv -Path $home\Desktop\ClientIP_Resolution.csv -NoTypeInformation -Encoding Unicode
($Data).count
