Add-PSSnapin *EXC*

Get-MessageTrackingLog -start 13:10 -eventID Receive | Where {$_.OriginalClientIP -NE $null} | Select-Object -ExpandProperty OriginalClientIP | Set-Content $home\desktop\ips.txt


$groupedLines = Get-Content "$home\desktop\ips.txt" | Group-Object

$uniqueLines = $groupedLines | ForEach-Object { "$($_.Name) - Appears $($_.Count) times" }
$uniqueLines | Set-Content "$home\desktop\unique.txt"
$uniqueLines
