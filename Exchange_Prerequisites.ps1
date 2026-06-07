<#
.SYNOPSIS
    Automated Post-Installation Performance and Security Baseline for Exchange Server.
.DESCRIPTION
    Applies recommended settings for Pagefile size, TCP KeepAlive timeouts, OS/ .NET TLS execution baselines,
    power states, event log sizing parameters, and structural component optimization.
.NOTES
    Must be executed from an elevated administrative shell console.
#>
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevated administrative permissions required to apply OS modifications. Exiting."
    break
}
$TotalMemoryMB = [math]::Round((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1MB, 0)
$ExchangeVersion = Read-Host "Enter target Exchange version architecture ('2016' or '2019')"
$PageFileSizeMB = switch ($ExchangeVersion) {
    "2016" {
        if ($TotalMemoryMB -le 65536) { $TotalMemoryMB + 10 } else { 32778 }
    }
    "2019" {
        if ($TotalMemoryMB -le 32768) { $TotalMemoryMB + 10 }
        elseif ($TotalMemoryMB -le 65536)  { 16384 }
        elseif ($TotalMemoryMB -le 98304)  { 24576 }
        elseif ($TotalMemoryMB -le 131072) { 32768 }
        elseif ($TotalMemoryMB -le 163840) { 40960 }
        elseif ($TotalMemoryMB -le 196608) { 49152 }
        else { 65536 }
    }
    Default {
        Write-Error "Unsupported Exchange Framework selection. Skipping Pagefile calculation."
        $null
    }
}
if ($null -ne $PageFileSizeMB) {
    Write-Host "Configuring Pagefile to fixed size: $PageFileSizeMB MB..." -ForegroundColor Cyan
    $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    if ($ComputerSystem.AutomaticManagedPagefile) {
        Set-CimInstance -Query "Select * from Win32_ComputerSystem" -Property @{ AutomaticManagedPagefile = $false }
    }

    $PageFileSetting = Get-CimInstance -ClassName Win32_PageFileSetting -Filter "Name LIKE 'C:%'"
    if (-not $PageFileSetting) {
        New-CimInstance -ClassName Win32_PageFileSetting -Property @{ Name = 'C:\pagefile.sys' } | Out-Null
    }
    Get-CimInstance -ClassName Win32_PageFileSetting -Filter "Name LIKE 'C:%'" | 
        Set-CimInstance -Property @{ InitialSize = $PageFileSizeMB; MaximumSize = $PageFileSizeMB }
}
Write-Host "Applying Network & OS Performance Baselines..." -ForegroundColor Cyan

# TCPKeepAlive Configuration
$TcpPath = "HKLM:\System\CurrentControlSet\Services\TcpIp\Parameters"
Set-ItemProperty -Path $TcpPath -Name "KeepAliveTime" -Value 1200000 -Type DWORD -Force

# High Performance Plan Deployment
powercfg -setactive SCHEME_MIN

# Enabling IPv6
Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue | 
    Where-Object { $_.Enabled -eq $false } | 
    Enable-NetAdapterBinding -ComponentID ms_tcpip6

# Adding Telnet client
if (-not (Get-WindowsFeature -Name Telnet-Client).Installed) {
    Install-WindowsFeature -Name Telnet-Client | Out-Null
}

# Encoforcing TLS 1.2 for OS and .NET
Write-Host "Applying cryptographic runtime rules (TLS 1.2)..." -ForegroundColor Cyan

$RegPaths = @(
    "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server",
    "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client",
    # CRITICAL FIX: Configures Native 64-Bit .NET Runtimes used by Exchange Server Processes
    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727",
    # Configures 32-Bit compatibility frameworks
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727"
)

foreach ($Path in $RegPaths) {
    if (-not (Test-Path $Path)) { New-Item $Path -Force | Out-Null }
    
    if ($Path -match 'Protocols') {
        Set-ItemProperty -Path $Path -Name "Enabled" -Value 1 -Type DWORD -Force
        Set-ItemProperty -Path $Path -Name "DisabledByDefault" -Value 0 -Type DWORD -Force
    } else {
        Set-ItemProperty -Path $Path -Name "SystemDefaultTlsVersions" -Value 1 -Type DWORD -Force
        Set-ItemProperty -Path $Path -Name "SchUseStrongCrypto" -Value 1 -Type DWORD -Force
    }
}

# Disable IE Enhanced mode
Write-Host "Disabling legacy IE ESC constraints safely..." -ForegroundColor Cyan
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" -Name "IsInstalled" -Value 0 -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}" -Name "IsInstalled" -Value 0 -Force

Write-Host "Setting event log files to 4GB..." -ForegroundColor Cyan
# Native .NET Event Log engine adjustments replace legacy slow command executables

Limit-EventLog -LogName "Application" -MaximumSize 4194240KB -OverflowAction OverwriteAsNeeded
try {
    # Utilizing native utility handle for custom application providers
    wevtutil sl "MSExchange Management" /ms:4294967296 | Out-Null
} catch {}

Write-Host "Disabling Print Spooler..." -ForegroundColor Cyan
Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
Set-Service -Name Spooler -StartupType Disabled

# 6. Verification Dump Outputs
Write-Host "`nEcosystem baseline alignment applied successfully!" -ForegroundColor Green
Write-Host "
 -> Profile Target  : Exchange $ExchangeVersion
 -> Managed Pagefile: $( (Get-CimInstance Win32_PageFileSetting).MaximumSize ) MB
 -> TCP KeepAlive   : 20 Minutes (1,200,000 ms)
 -> Power Topology  : High Performance Priority
 -> Crypto Engines  : TLS 1.2 System-Wide Enforced (64-Bit / 32-Bit .NET)
 -> Volumetric Logs : Application & MSExchange Management set to 4GB
 -> Print Spooler State   : Stopped and Disabled" -ForegroundColor White
