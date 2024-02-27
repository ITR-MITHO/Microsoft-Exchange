$Date = (Get-Date).AddDays(-60)
$Certificate = Get-ExchangeCertificate | Where {$_.NotAfter -lt "$Date"} | Select Thumbprint, FriendlyName, Subject, Notafter

CLS
Write-Host "Certificates that expires in less than 60 days.." -ForeGroundColor Yellow
Foreach ($Cert in $Certificate) {
if ($Certificate) {
$CertName = $Cert.Thumbprint
$CertFriendly = $Cert.friendlyname
$CertSubject = $Cert.Subject
$CertExpire = $Cert.NotAfter
