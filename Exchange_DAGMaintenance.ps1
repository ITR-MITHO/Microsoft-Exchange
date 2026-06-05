<#
.SYNOPSIS
    Orchestrates maintenance mode configurations for an Exchange Server DAG node.
.DESCRIPTION
    Runs locally on a target Exchange Mailbox server. Drains transport queues, 
    moves active database mounts, suspends cluster nodes, and gracefully brings 
    the node back into operational rotation.
.NOTES
    Must be executed from an elevated Exchange Management Shell session.
#>

# 1. Enforcement & Context Checks
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "This script requires administrative privileges. Please re-run from an elevated shell."
    break
}

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

$TargetServer = (Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop).Name

# 2. Functional Menu Controller (Replaces recursion with a structural loop)
while ($true) {
    Write-Host "`n=== Exchange Maintenance Orchestrator ===" -ForegroundColor Cyan
    Write-Host "Target Server: $TargetServer" -ForegroundColor DarkCyan
    Write-Host "1 - START MAINTENANCE MODE"
    Write-Host "2 - STOP MAINTENANCE MODE"
    Write-Host "3 - CHECK MAINTENANCE STATE"
    Write-Host "0 - EXIT"
    
    $Choice = Read-Host "`nPlease enter a selection"

    switch ($Choice) {
        "1" {
            Write-Host "`nCRITICAL WARNING:" -ForegroundColor Red
            Write-Host "Placing $TargetServer into maintenance mode will drain transport queues" -ForegroundColor Yellow
            Write-Host "and fail over all active databases. Connections will drop." -ForegroundColor Yellow
            
            $Confirm = Read-Host "Proceed with maintenance failover? (Y/N)"
            if ($Confirm -eq "Y") {
                Write-Host "Draining transport components..." -ForegroundColor Cyan
                Set-ServerComponentState -Identity $TargetServer -Component HubTransport -State Draining -Requester Maintenance
                
                # Dynamic targeted routing search
                $RedirectTarget = Get-ExchangeServer | Where-Object { $_.Name -ne $TargetServer -and ($_.IsHubTransportServer -or $_.IsMailboxServer) } | Select-Object -First 1
                if ($RedirectTarget) {
                    Write-Host "Redirecting active message queues to $($RedirectTarget.Fqdn)..." -ForegroundColor Cyan
                    Redirect-Message -Server $TargetServer -Target $RedirectTarget.Fqdn -Confirm:$false
                }

                Write-Host "Failing over active databases..." -ForegroundColor Cyan
                Move-ActiveMailboxDatabase -Server $TargetServer -SkipMoveSuppressionChecks -Confirm:$false > $null
                
                Write-Host "Pausing for cluster synchronization (15s)..." -ForegroundColor DarkCyan
                Start-Sleep -Seconds 15

                Write-Host "Suspending cluster node..." -ForegroundColor Cyan
                Suspend-ClusterNode -Name $TargetServer -ErrorAction SilentlyContinue
                
                Set-MailboxServer -Identity $TargetServer -DatabaseCopyActivationDisabledAndMoveNow $true -DatabaseCopyAutoActivationPolicy Blocked
                Set-ServerComponentState -Identity $TargetServer -Component ServerWideOffline -State Inactive -Requester Maintenance
                
                Write-Host "`nServer $TargetServer successfully isolated in Maintenance Mode." -ForegroundColor Green
            }
        }

        "2" {
            Write-Host "`nRestoring $TargetServer back to active production status..." -ForegroundColor Cyan
            
            Set-ServerComponentState -Identity $TargetServer -Component ServerWideOffline -State Active -Requester Maintenance
            Resume-ClusterNode -Name $TargetServer -ErrorAction SilentlyContinue
            Set-MailboxServer -Identity $TargetServer -DatabaseCopyAutoActivationPolicy Unrestricted -DatabaseCopyActivationDisabledAndMoveNow $false
            Set-ServerComponentState -Identity $TargetServer -Component HubTransport -State Active -Requester Maintenance

            Write-Host "Waiting for service convergence (10s)..." -ForegroundColor DarkCyan
            Start-Sleep -Seconds 10

            # Automated redistribution without breaking directory contexts
            $Dag = Get-DatabaseAvailabilityGroup | Where-Object { $_.Servers -contains $TargetServer } | Select-Object -First 1
            $ScriptPath = Join-Path $env:ExchangeInstallPath "Scripts\RedistributeActiveDatabases.ps1"
            
            if ($Dag -and (Test-Path $ScriptPath)) {
                Write-Host "Balancing DAG Active Databases across Cluster..." -ForegroundColor Cyan
                & $ScriptPath -DagName $Dag.Name -BalanceDbsByActivationPreference -SkipMoveSuppressionChecks -Confirm:$false
            } else {
                Write-Warning "Could not run database balancing script automatically. Verify DAG state."
            }

            Write-Host "`nServer $TargetServer is fully online and active." -ForegroundColor Green
        }

        "3" {
            Write-Host "`nRetrieving server component properties..." -ForegroundColor Cyan
            
            $HubState        = (Get-ServerComponentState -Identity $TargetServer -Component HubTransport).State
            $ServerWideState = (Get-ServerComponentState -Identity $TargetServer -Component ServerWideOffline).State
            $ClusterState    = (Get-ClusterNode -Name $TargetServer -ErrorAction SilentlyContinue).State
            $DbPolicy        = Get-MailboxServer -Identity $TargetServer | Select-Object DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow

            Write-Host "`n=== Current Component Breakdown ===" -ForegroundColor DarkCyan
            [PSCustomObject]@{
                "Hub Transport"       = $HubState
                "ServerWide Offline"  = $ServerWideState
                "Cluster Membership"  = $ClusterState
                "AutoActivation"      = $DbPolicy.DatabaseCopyAutoActivationPolicy
                "ActivationDisabled"  = $DbPolicy.DatabaseCopyActivationDisabledAndMoveNow
            } | Format-Table -AutoSize

            if ($DbPolicy.DatabaseCopyAutoActivationPolicy -eq "Unrestricted" -and 
                $DbPolicy.DatabaseCopyActivationDisabledAndMoveNow -eq $false -and 
                $HubState -eq "Active" -and 
                $ClusterState -eq "Up" -and 
                $ServerWideState -eq "Active") {
                Write-Host "STATUS: Node is healthy and fully operational in production." -ForegroundColor Green
            } else {
                Write-Host "STATUS: Node is currently in a MAINTENANCE or DRAINING state." -ForegroundColor Yellow
            }
        }

        "0" {
            Write-Host "Exiting orchestrator script." -ForegroundColor DarkYellow
            break
        }

        Default {
            Write-Warning "Invalid choice. Please pick an explicit menu option option [0-3]."
        }
    }
}
