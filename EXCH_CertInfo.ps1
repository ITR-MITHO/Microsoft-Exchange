<#
.SYNOPSIS
    Audits current usage of TLS certificates in an Exchange on-premises environment.

.DESCRIPTION
    Run this script before replacing a certificate to identify all components currently using it.
#>

# Ensure necessary snap-ins and modules are loaded
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.* -ErrorAction SilentlyContinue)) {
    try {
        Add-PSSnapin *Exchange* -ErrorAction Stop
    } catch {
        Write-Error "Failed to load Exchange snap-in. Ensure the Exchange Management Shell is installed."
        exit
    }
}

if (-not (Get-Module -Name ActiveDirectory)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Error "Failed to import Active Directory module. Ensure RSAT is installed."
        exit
    }
}

# Output current TLS certificate usage
Write-Host "`n=== Receive Connectors Using TLS Certificates ===" -ForegroundColor Cyan
Get-ReceiveConnector | Where-Object { $_.TlsCertificateName } | 
    Select-Object Identity, Enabled, TlsCertificateName, FQDN | 
    Format-Table -AutoSize

Write-Host "`n=== Send Connectors Using TLS Certificates ===" -ForegroundColor Cyan
Get-SendConnector | Where-Object { $_.TlsCertificateName } | 
    Select-Object Identity, Enabled, TlsCertificateName, FQDN | 
    Format-Table -AutoSize

# Count on-prem mailboxes
$mailboxCount = (Get-Mailbox -ResultSize Unlimited).Count
Write-Host "`n=== On-Prem Mailbox Count ===" -ForegroundColor Cyan
Write-Host "Total on-prem mailboxes: $mailboxCount"

# Identify the default SMTP certificate
Write-Host "`n=== Default SMTP Certificate ===" -ForegroundColor Cyan

try {
    $serverDN = (Get-ExchangeServer -Identity $env:COMPUTERNAME).DistinguishedName
    $adObject = Get-ADObject -Identity $serverDN -Properties msExchServerInternalTLSCert

    if ($adObject.msExchServerInternalTLSCert) {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import([Convert]::FromBase64String([Convert]::ToBase64String($adObject.msExchServerInternalTLSCert)))

        Write-Host "Thumbprint : $($cert.Thumbprint)"
        Write-Host "Friendly   : $($cert.FriendlyName)"
        Write-Host "Issuer     : $($cert.Issuer)"
        Write-Host "Subject    : $($cert.Subject)"
    } else {
        Write-Warning "No internal TLS certificate found in AD for this server."
    }
} catch {
    Write-Error "Failed to retrieve or parse the internal TLS certificate: $_"
}
