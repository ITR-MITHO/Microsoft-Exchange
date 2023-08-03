Get-ReceiveConnector | FL Identity
$Connector1 = Read-Host "Name of the reference connector"
$Connector2 = Read-Host "Name of the difference connector"

$Range1 = (Get-ReceiveConnector $Connector1).RemoteIPRanges | Sort-Object | Select Expression
$Range2 = (Get-ReceiveConnector $Connector2).RemoteIPRanges | Sort-Object | Select Expression

$Compare = Compare-Object -ReferenceObject ($Range1).Expression -DifferenceObject ($Range2).Expression
If ($Compare.SideIndicator -EQ "<=")
{
$Difference = ($Compare | Where SideIndicator -EQ "<=").InputObject
Echo "

$Connector2 is missing the following IP's from $Connector1

"
$Difference
}

$Compare2 = Compare-Object -ReferenceObject ($Range1).Expression -DifferenceObject ($Range2).Expression
If ($Compare2.SideIndicator -EQ "=>")
{
$Reference = ($Compare2 | Where SideIndicator -EQ "=>").InputObject
Echo "

$Connector1 is missing the following IP's from $Connector2

"
$Reference
}
