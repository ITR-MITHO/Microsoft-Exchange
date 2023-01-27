Get-DatabaseAvailabilityGroup | Fl Name
$EXCH01 = Read-Host "FQDN of the server that should be taken out of DAGMaintenence mode (eg. server.domain.com)"
$DAG = Read-Host "Name of Exchange DAG (Database Availability group)"

Set-ServerComponentState "$EXCH01" -Component ServerWideOffline -State Active -Requester Maintenance
Resume-ClusterNode -Name "$EXCH01"
Set-MailboxServer "$EXCH01" -DatabaseCopyAutoActivationPolicy Unrestricted
Set-MailboxServer "$EXCH01" -DatabaseCopyActivationDisabledAndMoveNow $false
Set-ServerComponentState "$EXCH01" -Component HubTransport -State Active -Requester Maintenance

Write-Host "Redistributing databases" -ForegroundColor Green

cd $exscripts
.\RedistributeActiveDatabases.ps1 -DagName "$DAG" -BalanceDbsByActivationPreference -SkipMoveSuppressionChecks -Confirm:$false

Write-Host "Server is now taken out of maintanence mode and databases redistributed." -ForegroundColor Green
