<#
.SYNOPSIS
    Sets Exchange and related services to default startup values after a patch installation.
#>
[CmdletBinding()]
param ()

Write-Host "Configuring Exchange services to default startup values..." -ForegroundColor Green

# Optimized: Provider-level filtering is significantly faster than pipeline filtering.
$ExServices = Get-Service -DisplayName "*Microsoft Exchange*"

foreach ($ExService in $ExServices) {
    if ($ExService.Name -eq "MSExchangeDiagnostics") {
        Set-Service -Name $ExService.Name -StartupType Automatic
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\MSExchangeDiagnostics" -Name "DelayedAutostart" -Value 1
    } else {
        Set-Service -Name $ExService.Name -StartupType Automatic
    }
}
# Optimized: Grouped into an array for maintainability. 
$RelatedServices = @(
    "FMS", "pla", "HealthService", "IISAdmin", 
    "SearchExchangeTracing", "W3Svc", "WinMgmt", "RemoteRegistry"
)

# Added resilience: Checks if the service exists before setting it. 
foreach ($Service in $RelatedServices) {
    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
        Set-Service -Name $Service -StartupType Automatic
    }
}

Write-Host "Exchange services configured to default startup values." -ForegroundColor Green
Write-Host "NOTE: POP and IMAP services might need their startup type adjusted." -ForegroundColor Yellow
Write-Host "Check C:\ExchangeSetupLogs\ServiceStartupMode.xml for pre-patch states (MSExchangePop3, MSExchangePOP3BE, MSExchangeImap4, MSExchangeIMAP4BE)." -ForegroundColor Yellow
