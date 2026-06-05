<#
.SYNOPSIS
    Orchestrates state management for local Microsoft Exchange ecosystem services.
.DESCRIPTION
    Runs locally on an on-premises Exchange Server. Captures baseline states of 
    active workloads, handles orderly service dependency teardowns, and restores 
    only pre-authorized workload configurations.
.OUTPUTS
    $Home\Desktop\ExchangeServices.csv - Captured baseline of active services.
#>

# 1. Access Control Enforcement
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevated administrative privileges required to manage system service states."
    break
}

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

$TargetServer = $env:COMPUTERNAME
$CsvPath      = Join-Path $home "Desktop\ExchangeServices.csv"

# 2. Dynamic Component Control Loop Engine
while ($true) {
    Write-Host "`n=== Exchange Service State Manager ===" -ForegroundColor Cyan
    Write-Host "Target Host: $TargetServer" -ForegroundColor DarkCyan
    Write-Host "1 - STOP AND DISABLE CURRENTLY RUNNING SERVICES"
    Write-Host "2 - RESTORE AND START BASELINE WORKLOADS"
    Write-Host "0 - EXIT"

    $Choice = Read-Host "`nPlease enter a selection"

    switch ($Choice) {
        "1" {
            # Capture active workloads *before* changing configuration layout definitions
            Write-Host "Evaluating current running service baseline..." -ForegroundColor Cyan
            
            # Captures all native MSExchange services + critical dependent Web infrastructure
            $ActiveServices = Get-Service | Where-Object { 
                ($_.Name -match '^MSExchange' -or $_.Name -in @('IISADMIN', 'W3SVC')) -and 
                $_.Status -eq 'Running' 
            }

            if ($ActiveServices.Count -eq 0 -and (-not (Test-Path $CsvPath))) {
                Write-Warning "No active Exchange services detected to baseline. Is the node already offline?"
                continue
            }

            # Only write the snapshot file if we have active services to baseline
            if ($ActiveServices.Count -gt 0) {
                $ActiveServices | Select-Object Name, DisplayName | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding Unicode -Force
                Write-Host "Active service state snapshot compiled to: $CsvPath" -ForegroundColor Green
            }

            Write-Host "`nWARNING: Stopping all Exchange capabilities on $TargetServer." -ForegroundColor Red
            $Confirm = Read-Host "Commit configuration state change? (Y/N)"
            
            if ($Confirm -eq "Y") {
                Write-Host "Initiating downstream workload teardown..." -ForegroundColor Cyan
                
                # Fetch fresh configurations to process
                $TargetList = Import-Csv -Path $CsvPath -ErrorAction SilentlyContinue
                
                foreach ($Svc in $TargetList) {
                    try {
                        # Force parameter stops dependents down the execution line automatically
                        Write-Host "Stopping and disabling service: $($Svc.Name)" -ForegroundColor DarkCyan
                        Stop-Service -Name $Svc.Name -Force -ErrorAction Stop
                        Set-Service -Name $Svc.Name -StartupType Disabled -ErrorAction Stop
                    } catch {
                        Write-Warning "Failed to cleanly transition state for object $($Svc.Name): $_"
                    }
                }
                Write-Host "Service teardown phase finalized successfully." -ForegroundColor Green
            }
        }

        "2" {
            if (-not (Test-Path $CsvPath)) {
                Write-Error "Baseline file missing at: $CsvPath. Cannot restore customized states safely."
                continue
            }

            Write-Host "Re-initializing core ecosystem workloads from baseline profile..." -ForegroundColor Cyan
            $BaselineServices = Import-Csv -Path $CsvPath

            # Step 1: Flip startup mappings back to automatic execution states
            foreach ($Svc in $BaselineServices) {
                try {
                    Set-Service -Name $Svc.Name -StartupType Automatic -ErrorAction SilentlyContinue
                } catch {}
            }

            # Step 2: Orderly startup sequence handling
            # MSExchangeADTopology must spin up fully first, or secondary workloads fail immediately
            if ($BaselineServices.Name -contains "MSExchangeADTopology") {
                Write-Host "Prioritizing Active Directory Topology Engine..." -ForegroundColor Yellow
                Start-Service -Name "MSExchangeADTopology" -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 5
            }

            foreach ($Svc in $BaselineServices) {
                if ((Get-Service -Name $Svc.Name).Status -eq 'Running') { continue }
                
                try {
                    Write-Host "Starting component workload: $($Svc.Name)..." -ForegroundColor DarkCyan
                    Start-Service -Name $Svc.Name -ErrorAction Stop
                } catch {
                    Write-Warning "Service startup timeout or failure on component: $($Svc.Name)"
                }
            }

            Write-Host "`nAll targeted baseline services received ignition instructions." -ForegroundColor Green
            Write-Host "RECOMMENDATION: Monitor Event Viewer or reboot node to guarantee initialization alignment." -ForegroundColor Yellow
        }

        "0" {
            Write-Host "Terminating execution routine." -ForegroundColor DarkYellow
            break
        }

        Default {
            Write-Warning "Invalid option target selected. Please choose [0-2]."
        }
    }
}
