<#
.SYNOPSIS
    Analyzes and configures TLS protocols and .NET strong crypto via Schannel registry keys.
#>

function Invoke-TlsManager {
    $protocols = @("TLS 1.0", "TLS 1.1", "TLS 1.2", "TLS 1.3")

    # Internal helper to test registry values
    function Test-RegistryValue {
        param([string]$Path, [string]$Name, [int]$Expected)
        if (Test-Path $Path) {
            $val = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue).$Name
            if ($null -eq $val) { 
                Write-Host "  [-] $Path\$Name - Missing" -ForegroundColor DarkYellow 
            } elseif ($val -ne $Expected) { 
                Write-Host "  [!] $Path\$Name - Incorrect (Expected: $Expected, Found: $val)" -ForegroundColor Yellow 
            } else { 
                Write-Host "  [+] $Path\$Name - OK" -ForegroundColor Green 
            }
        } else {
            Write-Host "  [-] $Path - Registry path missing" -ForegroundColor Red
        }
    }

    # Internal helper to set/create registry values
    function Set-RegistryValue {
        param([string]$Path, [string]$Name, [int]$Value)
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType 'DWord' -Force | Out-Null
    }

    # Main Execution Loop
    do {
        Write-Host "`n=================================" -ForegroundColor Cyan
        Write-Host "       TLS Management Tool       " -ForegroundColor Cyan
        Write-Host "=================================" -ForegroundColor Cyan
        Write-Host "1 - Check TLS Configuration"
        Write-Host "2 - Change TLS Configuration"
        Write-Host "0 - Exit"
        $choice = Read-Host "`nEnter selection"

        switch ($choice) {
            "1" {
                # Check protocols
                foreach ($proto in $protocols) {
                    Write-Host "`nChecking $proto..." -ForegroundColor Cyan
                    Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server" -Name "Enabled" -Expected 1
                    Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server" -Name "DisabledByDefault" -Expected 0
                    Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Client" -Name "Enabled" -Expected 1
                    Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Client" -Name "DisabledByDefault" -Expected 0
                }
                
                # Check .NET keys
                Write-Host "`nChecking .NET Framework Strong Crypto Keys..." -ForegroundColor Cyan
                $netPaths = @(
                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319",
                    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727",
                    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727"
                )
                foreach ($path in $netPaths) {
                    Test-RegistryValue -Path $path -Name "SystemDefaultTlsVersions" -Expected 1
                    Test-RegistryValue -Path $path -Name "SchUseStrongCrypto" -Expected 1
                }
            }
            "2" {
                # Configure protocols
                foreach ($proto in $protocols) {
                    Write-Host "`n[ Target: $proto ]" -ForegroundColor Magenta
                    $action = Read-Host "Action (E=Enable / D=Disable / S=Skip)"
                    
                    if ($action -match '^[EDed]$') {
                        $enabled = if ($action -match '^[Ee]$') { 1 } else { 0 }
                        $disabledByDefault = if ($action -match '^[Ee]$') { 0 } else { 1 }

                        Set-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server" -Name "Enabled" -Value $enabled
                        Set-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server" -Name "DisabledByDefault" -Value $disabledByDefault
                        Set-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Client" -Name "Enabled" -Value $enabled
                        Set-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Client" -Name "DisabledByDefault" -Value $disabledByDefault
                        
                        Write-Host ">> $proto configured to $action" -ForegroundColor Green
                    } elseif ($action -match '^[Ss]$') {
                        Write-Host ">> Skipped $proto" -ForegroundColor DarkGray
                    }
                }

                # Configure .NET keys
                Write-Host "`n[ Target: .NET Framework Strong Crypto ]" -ForegroundColor Magenta
                $netAction = Read-Host "Action (E=Enable / D=Disable / S=Skip)"
                
                if ($netAction -match '^[EDed]$') {
                    $val = if ($netAction -match '^[Ee]$') { 1 } else { 0 }
                    $netPaths = @(
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319",
                        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727",
                        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727"
                    )
                    foreach ($path in $netPaths) {
                        Set-RegistryValue -Path $path -Name "SystemDefaultTlsVersions" -Value $val
                        Set-RegistryValue -Path $path -Name "SchUseStrongCrypto" -Value $val
                    }
                    Write-Host ">> .NET Framework keys configured to $netAction" -ForegroundColor Green
                } elseif ($netAction -match '^[Ss]$') {
                    Write-Host ">> Skipped .NET Framework Strong Crypto" -ForegroundColor DarkGray
                }

                Write-Host "`n[!] IMPORTANT: A reboot is required to apply Schannel changes." -ForegroundColor Red
            }
            "0" { 
                Write-Host "Exiting..."
                break 
            }
            default { 
                Write-Host "Invalid selection." -ForegroundColor Red 
            }
        }
    } while ($true)
}

# Execute the function
Invoke-TlsManager
