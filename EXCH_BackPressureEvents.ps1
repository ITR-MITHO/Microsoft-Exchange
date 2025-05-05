<#
Searches for Event ID 15004, 15005, 15006 and 15007 in the Application log
https://learn.microsoft.com/en-us/exchange/mail-flow/back-pressure?view=exchserver-2019#back-pressure-logging-information

#>
$eventIDs = @(15004, 15005, 15006, 15007)
$startTime = (Get-Date).AddHours(-24)
$events = Get-WinEvent -FilterHashtable @{
    LogName = 'Application';
    ProviderName = 'MSExchangeTransport';
    StartTime = $startTime;
    ID = $eventIDs
} -ComputerName $env:computername -ErrorAction SilentlyContinue


if ($events.Count -eq 0) {
    Write-Host "No backpressure events found in the last 24 hours on $server."
} else {
    Write-Host "Backpressure events found on $server in the last 24 hours:"
    
    # Display the events
    foreach ($event in $events) {
        Write-Host "Time: $($event.TimeCreated)"
        Write-Host "Event ID: $($event.Id)"
        Write-Host "Message: $($event.Message)"
        Write-Host "---------------------------------------------"
    }
}
