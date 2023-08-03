Get-ReceiveConnector | FL Identity
$Connector1 = Read-Host "Enter Reference connector name"
$Connector2 = Read-Host "Enter Difference connector name"

$Range1 = (Get-ReceiveConnector $Connector1).RemoteIPRanges | Sort-Object | Select Expression
$Name1 = $Range1
$Range2 = (Get-ReceiveConnector $Connector2).RemoteIPRanges | Sort-Object | Select Expression
$Name2 = $Range2

Write-Host "=> Means the value can only be found in the Difference object" -ForegroundColor Yellow
Write-Host "<= Means the value can only be found in the Reference object" -ForegroundColor Yellow
Compare-Object -ReferenceObject ($Name1).Expression -DifferenceObject ($Name2).Expression
