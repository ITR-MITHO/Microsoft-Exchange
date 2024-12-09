<#

Change the $ExportPath to your designated path for the output file.
Script will prompt for Exchange Administrator/Global Administrator credentials
For each accepted domain in Exchange it will check DNS records

#>

Connect-ExchangeOnline

$ExportPath = "C:\Users\mitho\Desktop"
$ErrorActionPreference = 'SilentlyContinue'
$Domains = Get-AcceptedDomain | Select DomainName
$Result = foreach ($Domain in $Domains) {
$DomainName = $Domain.DomainName

# MX Records
$MX = nslookup -q=mx $DomainName 8.8.8.8 2>$null
$MXGrabber = $MX | Select-String "mail exchanger"

# SPF Records
$SPF = nslookup -q=txt $DomainName 8.8.8.8 2>$null
$SPFGrabber = $SPF | Select-String "spf"

# DMARC Records
$DMARC = nslookup -q=txt _dmarc.$DomainName 8.8.8.8 2>$null
$DMARCGrabber = $DMARC | Select-String "v=DMARC1"

# Random SPF Record
$RandomSPF = nslookup -q=txt randomspfrecord.$DomainName 8.8.8.8 2>$null
$RandomSPFGrabber = $RandomSPF | Select-String "v=spf1"

# Selector1 CNAME Record
$Selector1 = nslookup -q=cname "selector1._domainkey.$DomainName" 8.8.8.8 2>$null
$Selector1Grabber = $Selector1 | Select-String "canonical name"

# Selector2 CNAME Record
$Selector2 = nslookup -q=cname "selector2._domainkey.$DomainName" 8.8.8.8 2>$null
$Selector2Grabber = $Selector2 | Select-String "canonical name"

# Display Results
Echo "--- DNS Information for $DomainName ---" >> "$ExportPath\DomainChecker.txt"

# MX Record
If ($MXGrabber) {
    $MXRecord = $MXGrabber.Line -replace ".*mail exchanger = ", ""
    Echo "MX-Record:
$MXRecord
    " >> "$ExportPath\DomainChecker.txt"
} Else {
    Echo "MX-record:
No valid MX-record found for $DomainName
    " >> "$ExportPath\DomainChecker.txt"
}

# SPF Record
If ($SPFGrabber) {
    $SPFRecord = $SPFGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
    Echo "SPF-record:
$SPFRecord
    " >> "$ExportPath\DomainChecker.txt"
    
} Else {
    Echo "SPF-record:
No valid SPF-record found for $DomainName
    " >> "$ExportPath\DomainChecker.txt"
}

# Random SPF Record
If ($RandomSPFGrabber) {
    $RandomSPF = $RandomSPFGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
    Echo "Random SPF:
$RandomSPF
    " >> "$ExportPath\DomainChecker.txt"
    
} Else {
    Echo "Random SPF:
No Random SPF found for $DomainName
    " >> "$ExportPath\DomainChecker.txt"
}


# DMARC Record
If ($DMARCGrabber) {
    $DMARCRecord = $DMARCGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
 Echo "DMARC-record:
$DMARCRecord
    " >> "$ExportPath\DomainChecker.txt"
    
} Else {
    Echo "DMARC-record:
No valid DMARC-record found for $DomainName
    " >> "$ExportPath\DomainChecker.txt"
}

# Selector1 Record
If ($Selector1Grabber) {
$S1 = $Selector1Grabber.Line -replace ".*canonical name = ", ""
 Echo "Selector1:
$S1
    " >> "$ExportPath\DomainChecker.txt"
    
} Else {
    Echo "Selector1:
No CNAME found for Selector1 $DomainName
    " >> "$ExportPath\DomainChecker.txt"
}


# Selector2 Rcord
If ($Selector2Grabber) {
$S2 = $Selector2Grabber.Line -replace ".*canonical name = ", ""
 Echo "Selector2:
$S2
    
    
    
    " >> "$ExportPath\DomainChecker.txt"
    
} Else {
    Echo "Selector2:
No CNAME found for Selector2 $DomainName
    
    
    
    
    " >> "$ExportPath\DomainChecker.txt"
}
    }
Write-Host "Export completed, find your file here; $ExportPath\DomainCheckter.txt" -ForeGroundColor Green
