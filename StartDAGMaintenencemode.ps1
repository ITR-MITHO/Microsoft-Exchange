$EXCH01 = Read-Host "FQDN of the server that should be in DAGMaintenence mode (eg. server.domain.com)"
$EXCH02 = Read-Host "FQDN of the server that should be active during patching (eg. server.domain.com)"

Write-Host "Draining message queue...." -ForegroundColor Green
Set-ServerComponentState -Identity "$EXCH01" -Component HubTransport -State Draining -Requester Maintenance
cls
Write-Host "Message queue drained!" -ForegroundColor Green

Redirect-Message -Server "$EXCH01" -Target "$EXCH02" -Confirm:$false
Suspend-ClusterNode "$EXCH01"
Set-MailboxServer "$EXCH01" -DatabaseCopyActivationDisabledAndMoveNow $true
Set-MailboxServer "$EXCH01" -DatabaseCopyAutoActivationPolicy Blocked

Set-ServerComponentState "$EXCH01" -Component ServerWideOffline -State Inactive -Requester Maintenance

cls
Write-Host "If the servername is shown here below, the server is in maintenence mode." -ForegroundColor Green
Get-databaseavailabilitygroup -status | fl name,ServersInMaintenance 