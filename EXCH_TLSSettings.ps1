<#
First the script will prompt if you want to check current TLS settings or change TLS Settings

The script will prompt you for each of the TLS protocols; 1.0, 1.1 and 1.2
E = Enable
D = Disable

#>

Function TLS { 
$Function = Read-Host "
Please enter one of the below numbers to proceed:

    1 - CHECK TLS CONFIGURATION
    2 - CHANGE TLS CONFIGURATION
    0 - EXIT
    "

If ($Function -EQ "1")
    {

# TLS 1.2
# Define registry keys and expected values
$registryKeys = @(
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"; Name="Enabled"; Expected=1},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"; Name="DisabledByDefault"; Expected=0},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"; Name="Enabled"; Expected=1},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"; Name="DisabledByDefault"; Expected=0},
    @{Path="HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"; Name="SystemDefaultTlsVersions"; Expected=1},
    @{Path="HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"; Name="SchUseStrongCrypto"; Expected=1}
)

Write-Host "Checking TLS 1.2 Registry Keys..." -ForegroundColor Cyan

foreach ($entry in $registryKeys) {
    $path = $entry.Path
    $name = $entry.Name
    $expectedValue = $entry.Expected

    if (Test-Path $path) {
        $actualValue = (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue).$name

        if ($null -eq $actualValue) {
            Write-Host "$path\$name - Missing" -ForegroundColor Yellow
        }
        elseif ($actualValue -ne $expectedValue) {
            Write-Host "$path\$name - Incorrect Value (Expected: $expectedValue, Found: $actualValue)" -ForegroundColor Yellow
        }
        else {
            Write-Host "$path\$name - OK" -ForegroundColor Green
        }
    }
    else {
        Write-Host "$path - Registry path missing" -ForegroundColor Red
    }
}

# TLS 1.1

# Define registry keys and expected values
$registryKeys = @(
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server"; Name="Enabled"; Expected=1},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server"; Name="DisabledByDefault"; Expected=0},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client"; Name="Enabled"; Expected=1},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client"; Name="DisabledByDefault"; Expected=0}
)

Write-Host "Checking TLS 1.1 Registry Keys..." -ForegroundColor Cyan

foreach ($entry in $registryKeys) {
    $path = $entry.Path
    $name = $entry.Name
    $expectedValue = $entry.Expected

    if (Test-Path $path) {
        $actualValue = (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue).$name

        if ($null -eq $actualValue) {
            Write-Host "$path\$name - Missing" -ForegroundColor Yellow
        }
        elseif ($actualValue -ne $expectedValue) {
            Write-Host "$path\$name - Incorrect Value (Expected: $expectedValue, Found: $actualValue)" -ForegroundColor Yellow
        }
        else {
            Write-Host "$path\$name - OK" -ForegroundColor Green
        }
    }
    else {
        Write-Host "$path - Registry path missing" -ForegroundColor Red
    }
}

# TLS 1.0

# Define registry keys and expected values
$registryKeys = @(
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server"; Name="Enabled"; Expected=1},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server"; Name="DisabledByDefault"; Expected=0},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client"; Name="Enabled"; Expected=1},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client"; Name="DisabledByDefault"; Expected=0}
)

Write-Host "Checking TLS 1.0 Registry Keys..." -ForegroundColor Cyan

foreach ($entry in $registryKeys) {
    $path = $entry.Path
    $name = $entry.Name
    $expectedValue = $entry.Expected

    if (Test-Path $path) {
        $actualValue = (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue).$name

        if ($null -eq $actualValue) {
            Write-Host "$path\$name - Missing" -ForegroundColor Yellow
        }
        elseif ($actualValue -ne $expectedValue) {
            Write-Host "$path\$name - Incorrect Value (Expected: $expectedValue, Found: $actualValue)" -ForegroundColor Yellow
        }
        else {
            Write-Host "$path\$name - OK" -ForegroundColor Green
        }
    }
    else {
        Write-Host "$path - Registry path missing" -ForegroundColor Red
    }
}
    }

If ($Function -EQ "2")
    {

# TLS 1.2 Settings
$TLS12 = Read-Host "Enable or Disable TLS 1.2? (E/D)"
If ($TLS12 -eq "E")
{
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



If (-Not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319')) {
    New-Item 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SystemDefaultTlsVersions' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -value '1' -PropertyType 'DWord' -Force | Out-Null


If (-Not (Test-Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727')) {
    New-Item 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727' -Name 'SystemDefaultTlsVersions' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727' -Name 'SchUseStrongCrypto' -value '1' -PropertyType 'DWord' -Force | Out-Null
Write-Host "Enabled TLS 1.2 for OS and .NET" -ForegroundColor Green
}

If ($TLS12 -eq "D")
{
If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server')) {
    New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Name 'Enabled' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Name 'DisabledByDefault' -value 1 -PropertyType 'DWord' -Force | Out-Null

If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Name 'Enabled' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Name 'DisabledByDefault' -value 1 -PropertyType 'DWord' -Force | Out-Null


If (-Not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319')) {
    New-Item 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SystemDefaultTlsVersions' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -value '0' -PropertyType 'DWord' -Force | Out-Null

If (-Not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727')) {
    New-Item 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Name 'SystemDefaultTlsVersions' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Name 'SchUseStrongCrypto' -value '0' -PropertyType 'DWord' -Force | Out-Null

Write-Host "Disabled TLS 1.2 for OS and .NET" -ForegroundColor Yellow
}

# TLS 1.1 Settings
$TLS11 = Read-Host "Enable or Disable TLS 1.1? (E/D)"
If ($TLS11 -EQ "D")
{
If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'Enabled' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'DisabledByDefault' -value 1 -PropertyType 'DWord' -Force | Out-Null

If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Name 'Enabled' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Name 'DisabledByDefault' -value 1 -PropertyType 'DWord' -Force | Out-Null

Write-Host "TLS 1.1 has been disabled" -ForegroundColor Yellow
}

If ($TLS11 -EQ "E")
{

If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'Enabled' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'DisabledByDefault' -value 0 -PropertyType 'DWord' -Force | Out-Null

If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Name 'Enabled' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Name 'DisabledByDefault' -value 0 -PropertyType 'DWord' -Force | Out-Null
Write-Host "TLS 1.1 has been enabled" -ForegroundColor Green
}


 # TLS 1.0 Settings
$TLS10 = Read-Host "Enable or Disable TLS 1.0? (E/D)"
If ($TLS10 -EQ "D")
{
If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Name 'Enabled' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Name 'DisabledByDefault' -value 1 -PropertyType 'DWord' -Force | Out-Null

If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -Name 'Enabled' -value '0' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -Name 'DisabledByDefault' -value 1 -PropertyType 'DWord' -Force | Out-Null
Write-Host 'TLS 1.0 has been disabled.' -ForegroundColor Yellow
}

If ($TLS10 -EQ "E")
{
If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Name 'Enabled' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Name 'DisabledByDefault' -value 0 -PropertyType 'DWord' -Force | Out-Null

If (-Not (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client')) {
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -Force | Out-Null
}
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -Name 'Enabled' -value '1' -PropertyType 'DWord' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -Name 'DisabledByDefault' -value 0 -PropertyType 'DWord' -Force | Out-Null
Write-Host 'TLS 1.0 has been enabled.' -ForegroundColor Green
}

Write-Host "
IMPORTANT: You need to reboot before the changes are applied" -ForeGroundColor Red
}

If ($Function -EQ "0")
{  
    Exit
}
    } 

TLS
