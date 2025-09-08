<#

The script is designed to set all Exchange services to the default startup value, in case a patch disables them all.

#>

Write-Host "Configuring Exchange services to default startup values." -ForegroundColor Green

$ExServices = Get-Service | ?{$_.DisplayName -like "*Microsoft Exchange*"}

foreach ($ExService in $ExServices)
{
    if ($ExService.Name -eq "MSExchangeDiagnostics")    
    {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\MSExchangeDiagnostics" -Name "Start" -Value 2
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\MSExchangeDiagnostics" -Name "DelayedAutostart" -Value 1
    }
    else
    {
        Set-Service -Name $ExService.Name -StartupType Automatic
    }
}

#Configuring startup related servies

# Microsoft Filtering Management Service
Set-Service -Name FMS -StartupType Automatic

# Performance Logs & Alerts
Set-Service -Name pla -StartupType Automatic

# Microsoft Monitoring Agent
Set-Service -Name HealthService -StartupType Automatic

# IIS Admin Service
Set-Service -Name IISAdmin -StartupType Automatic

# Tracing Service for Search in Exchange
Set-Service -Name SearchExchangeTracing -StartupType Automatic

# World Wide Web Publishing Service
Set-Service -Name W3Svc -StartupType Automatic

# Windows Management Instrumentation
Set-Service -Name WinMgmt -StartupType Automatic

# Remote Registry (should be "Automatic - Trigger start" when viewed afterwards
Set-Service -Name RemoteRegistry -StartupType Automatic

Write-Host "Exchange services configured to default startup values." -ForegroundColor Green
Write-Host "POP and IMAP services might need to have the startup type changed. Check the ExchangeSetuplog folder for the pre-patch setup for the following services: MSExchangePop3, MSExchangePOP3BE, MSExchangeImap4 and MSExchangeIMAP4BE and change the startup type manually if needed." -ForegroundColor Yellow
Write-Host "Find the info here: C:\ExchangeSetupLogs\ServiceStartupMode.xml" -ForegroundColor Yellow
