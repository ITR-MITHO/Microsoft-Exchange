Get-ReceiveConnector | FL Identity
$Connector1 = Read-Host "Name of the reference connector"
$Connector2 = Read-Host "Name of the difference connector"

$Range1 = (Get-ReceiveConnector $Connector1).RemoteIPRanges | Sort-Object | Select Expression
$Range2 = (Get-ReceiveConnector $Connector2).RemoteIPRanges | Sort-Object | Select Expression


Write-Host "=> Means the value can only be found in the Difference object" -ForegroundColor Yellow
Write-Host "<= Means the value can only be found in the Reference object" -ForegroundColor Yellow
$Compare = Compare-Object -ReferenceObject ($Range1).Expression -DifferenceObject ($Range2).Expression

If ($Compare.SideIndicator -EQ "<=")

{

$Difference = $Compare | Where SideIndicator -EQ "<=" | Select InputObject

Write-Host "

$Connector2 is missing the following IP's from $Connector1" -ForegroundColor Red
$Difference

}
