$Date = (Get-Date).AddDays(-60)
$Certificate = Get-ExchangeCertificate | Where {$_.NotAfter -lt "$Date"} | Select Thumbprint, FriendlyName, Subject, Notafter

Foreach ($Cert in $Certificate) {
if ($Certificate) {
$CertName = $Cert.Thumbprint
$CertFriendly = $Cert.friendlyname
$CertSubject = $Cert.Subject
$CertExpire = $Cert.NotAfter

# Change TO and FROM
Send-MailMessage -To MyEmail@Domain.com -From Exchange@domain1.com -Subject "Certificate expires" -Body "
 Certificate Name: $Certname
 Friendly Name: $CertFriendly
 Subject: $CertSubject
 Expiration date: $CertExpire" -SmtpServer localhost
}
}
