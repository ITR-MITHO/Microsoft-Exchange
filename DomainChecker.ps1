<#

The script will ask you to enter a domain name (e.g. domain.com) once entered it will look for the following records for the domain entered
MX-record
SPF-record
Random SPF
DMARC-record
DKIM selector 1 and DKIM Selector 2

#>

$Domain = Read-Host "Enter domain (e.g. domain.com)"

# MX Records
$MX = nslookup -q=mx $Domain 2>$null
$MXGrabber = $MX | Select-String "mail exchanger"

# SPF Records
$SPF = nslookup -q=txt $Domain 2>$null
$SPFGrabber = $SPF | Select-String "spf"

# DMARC Records
$DMARC = nslookup -q=txt "_dmarc.$Domain" 2>$null
$DMARCGrabber = $DMARC | Select-String "v=DMARC1"

# Random SPF Record
$RandomSPF = nslookup -q=txt "randomspfrecord.$Domain" 2>$null
$RandomSPFGrabber = $RandomSPF | Select-String "v=spf1"

# Selector1 CNAME Record
$Selector1 = nslookup -q=cname "selector1._domainkey.$Domain" 2>$null
$Selector1Grabber = $Selector1 | Select-String "canonical name"

# Selector2 CNAME Record
$Selector2 = nslookup -q=cname "selector2._domainkey.$Domain" 2>$null
$Selector2Grabber = $Selector2 | Select-String "canonical name"

# Display Results
Write-Host "`n--- DNS Information for $Domain ---" -ForegroundColor Yellow

# MX Record
If ($MXGrabber) {
    Write-Host "`nMX Record:" -ForegroundColor Yellow
    $MXGrabber.Line -replace ".*mail exchanger = ", ""
} Else {
    Write-Host "`nNo MX records found for $Domain"
}

# SPF Record
If ($SPFGrabber) {
    Write-Host "`nSPF Record:" -ForegroundColor Yellow
    $SPFGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
} Else {
    Write-Host "`nNo valid SPF record found for $Domain"
}

# Random SPF Record
If ($RandomSPFGrabber) {
    Write-Host "`nRandom SPF Record:" -ForegroundColor Yellow
    $RandomSPFGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
} Else {
    Write-Host "`nNo Random SPF record found for $Domain"
}

# DMARC Record
If ($DMARCGrabber) {
    Write-Host "`nDMARC Record:" -ForegroundColor Yellow
    $DMARCGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
} Else {
    Write-Host "`nNo valid DMARC record found for $Domain"
}

# Selector1 Record
If ($Selector1Grabber) {
    Write-Host "`nSelector1 Record:" -ForegroundColor Yellow
    $Selector1Grabber.Line -replace ".*canonical name = ", ""
} Else {
    Write-Host "`nNo valid Selector1 CNAME record found for $Domain"
}

# Selector2 Record
If ($Selector2Grabber) {
    Write-Host "`nSelector2 Record:" -ForegroundColor Yellow
    $Selector2Grabber.Line -replace ".*canonical name = ", ""
} Else {
    Write-Host "`nNo valid Selector2 CNAME record found for $Domain"
}
