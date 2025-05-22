<# 

Can be used before changing a certificate on-prem to show where it is currently used.

#>
Add-PSSnapin *EXC*
Import-Module ActiveDirectory
# Receive Connectors
Write-Host "Receive Connectors" 
Get-ReceiveConnector | Where {$_.TlsCertificateName -NE $null} | Select Identity, Enabled, TlsCertificateName, FQDN


# Send Connectors
Write-Host "
Send connectors"
Get-SendConnector | Where {$_.TlsCertificateName -NE $null} | Select Identity, Enabled, TlsCertificateName, FQDN

# Mailboxes on-prem
Write-Host "
Mailboxes on-prem)"
(Get-Mailbox).count

# Default SMTP Cert

$DNAME = (Get-ExchangeServer -Identity $env:computername).distinguishedname
$TransportCert = (Get-ADObject -Identity $DNAME -Properties *).msExchServerInternalTLSCert
$Cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$CertBlob = [System.Convert]::ToBase64String($TransportCert)
$Cert.Import([Convert]::FromBase64String($CertBlob))

$CertThumb = $Cert.Thumbprint
$CertName = $Cert.FriendlyName
$CertIssuer = $Cert.Issuer

Write-Host "
Default SMTP Certificate

Thumbprint: $CertThumb
Friendly: $CertName
Issuer: $CertIssuer

"
