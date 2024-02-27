$Date = (Get-Date).AddDays(-60)
$Certificate = Get-ExchangeCertificate | Where {$_.NotAfter -lt "$Date"} | Select Thumbprint, FriendlyName, Subject, Notafter, services

CLS
Write-Host "Certificates that expires in less than 60 days.." -ForeGroundColor Yellow
Foreach ($Cert in $Certificate) {
if ($Certificate) {
$Thumb = $Cert.Thumbprint
$Friendly = $Cert.friendlyname
$Subject = $Cert.Subject
$NotAfter = $Cert.NotAfter
$Services = $Cert.Services

Write-Host "
$Friendly
$Subject
$Thumb
$NotAfter
$Services

"
}
  }
