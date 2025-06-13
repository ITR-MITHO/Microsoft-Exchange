<#
Automatically configures the below settings: 

Pagefile based on installed RAM and Exchange version
TCPKeepAlive - 20 minutes
Power Plan: High Performance
Enables IPv6 on all NIC
Enables TLS 1.2 for OS and .NET
Install Windows Feature: Telnet client
Disables IE Enhanced Security
Application and MSExchange Management event log size set to 4GB
Stop and disable Print Spooler service
#>

# Elevated PowerShell Check
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
Write-Host "PowerShell needs to be elevated" -ForeGroundColor Yellow
Break
}

# Pagefile
$RAM = (Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1MB
function Set-PageFile-Exchange2013_2016 {
    if ($RAM -le 8192) {
        $PageFileSize = 8202
    } elseif ($RAM -le 16384) {
        $PageFileSize = 16394
    } elseif ($RAM -le 32768) {
        $PageFileSize = 32778
    } elseif ($RAM -le 65536) {
        $PageFileSize = 32778
    } elseif ($RAM -le 131072) {
        $PageFileSize = 32778
    } elseif ($RAM -le 196608) {
        $PageFileSize = 32778
    }
Set-PageFile -InitialSize $PageFileSize -MaximumSize $PageFileSize
}
function Set-PageFile-Exchange2019 {
    if ($RAM -le 16384) {
        $PageFileSize = 16394 
    } elseif ($RAM -le 32768) {
        $PageFileSize = 32778 
    } elseif ($RAM -le 65536) {
        $PageFileSize = 16384  
    } elseif ($RAM -le 98304) {
        $PageFileSize = 24576
    } elseif ($RAM -le 131072) {
        $PageFileSize = 32768
    } elseif ($RAM -le 163840) {
        $PageFileSize = 40960
    } elseif ($RAM -le 196608) {
        $PageFileSize = 49152
    } elseif ($RAM -le 262144) {
        $PageFileSize = 65536
    }
    Set-PageFile -InitialSize $PageFileSize -MaximumSize $PageFileSize
}
function Set-PageFile {
    param (
        [int]$InitialSize,
        [int]$MaximumSize
    )
    $pagefileSetting = Get-WmiObject -Query "Select * from Win32_ComputerSystem"
    $pagefileSetting.AutomaticManagedPagefile = $false
    $pagefileSetting.Put()

    $pagefile = Get-WmiObject -Query "Select * from Win32_PageFileSetting where Name = 'C:\\pagefile.sys'"
    if ($pagefile -eq $null) {
        # If the pagefile does not exist, create it
        $pagefile = ([WmiClass] "root\cimv2:Win32_PageFileSetting").CreateInstance()
        $pagefile.Name = "C:\\pagefile.sys"
    }
    $pagefile.InitialSize = $InitialSize
    $pagefile.MaximumSize = $MaximumSize
    $pagefile.Put()
}

$ExchangeVersion = Read-Host "Enter Exchange version - e.g. '2016' or '2019'"
if ($ExchangeVersion -EQ "2016") {
    Set-PageFile-Exchange2013_2016 | Out-Null
} elseif ($ExchangeVersion -EQ "2019") {
    Set-PageFile-Exchange2019 | Out-Null
} else {
    Write-Output "Unsupported Exchange version"
}

# TCPKeepAlive (20 minutes)
New-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\TcpIp\Parameters" -Name "KeepAliveTime" -PropertyType DWORD -Value 1200000 -Force | Out-Null

# High Performance Power Plan
powercfg -setactive SCHEME_MIN

# Enable IPV6 on all NIC
$NIC = Get-NetAdapterBinding -ComponentID ms_tcpip6 | Select Name
Foreach ($N in $NIC)

{
	$Name = $N.Name
	Enable-NetAdapterBinding -Name $Name -ComponentID ms_tcpip6
}

# Telnet Client
Install-WindowsFeature -Name Telnet-Client | Out-Null

# TLS Settings
# Enable TLS 1.2
If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server')) {
    New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Name 'Enabled' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Name 'DisabledByDefault' -value 0 -PropertyType 'DWord' -Force | Out-Null

If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Name 'Enabled' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Name 'DisabledByDefault' -value 0 -PropertyType 'DWord' -Force | Out-Null

# Enable TLS 1.2 for .NET 4.x
If (-Not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319')) {
    New-Item 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SystemDefaultTlsVersions' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -value '1' -PropertyType 'DWord' -Force | Out-Null

# Enable TLS 1.2 for .NET 3.5
If (-Not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727')) {
    New-Item 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Name 'SystemDefaultTlsVersions' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Name 'SchUseStrongCrypto' -value '1' -PropertyType 'DWord' -Force | Out-Null

# Disable IE Enhanced Security
$AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
$UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0 -Force
Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 0 -Force
Stop-Process -Name Explorer -Force

# Set Application and MSExchange Management event log size set to 4GB
wevtutil sl "MSExchange Management" /ms:4294967296
wevtutil sl "Application" /ms:4294967296


# Stop the Print Spooler service
Stop-Service -Name Spooler -Force
Set-Service -Name Spooler -StartupType Disabled


$PageSize = (Get-CimInstance Win32_PageFileSetting).MaximumSize
Write-Host "
Adjusted the following settings!" -ForegroundColor Yellow

Write-Host "

User input: Exchange $ExchangeVersion
PageFile: $PageSize MB
TCPKeepAlive: 20 minutes
Power Plan: High performance
Telnet Client: Installed
IPv6: ENABLED
TLS 1.2 Enabled (OS and .NET)
Application and MSExchange Management event log size set to 4GB
Disabled IE Enhanced Security
Stopped and disabled Print Spooler service" -ForegroundColor Green
