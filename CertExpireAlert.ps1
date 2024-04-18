$Date = (Get-Date).AddDays(-30)
$Certificate = Get-ExchangeCertificate | Where {$_.NotAfter -lt "$Date"} | Select Thumbprint, FriendlyName, Subject, Notafter, services
CLS
Write-Host "Certificates that expires in less than 30 days.." -ForeGroundColor Yellow
Foreach ($Cert in $Certificate) {
if ($Certificate) {
$Thumb = $Cert.Thumbprint
$Friendly = $Cert.friendlyname
$Subject = $Cert.Subject
$NotAfter = $Cert.NotAfter.ToString("dd-MM-yyyy")
$Services = $Cert.Services

Write-Host "
Friendly: $Friendly
Subject: $Subject
Thumbprint: $Thumb
Expires: $NotAfter
Services: $Services
"
}
  }
