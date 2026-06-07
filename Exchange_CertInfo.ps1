<#
.SYNOPSIS
    Audits current usage of TLS certificates in an Exchange on-premises environment.
.DESCRIPTION
    Run this script before replacing a certificate to identify all connectors and default SMTP certificates.
.OUTPUTS
    Outputs tables directly to the console host.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

Write-Host "`n=== Mailbox Inventory Counts ===" -ForegroundColor Cyan
$OnPremCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox -ErrorAction SilentlyContinue).Count
$RemoteCount = (Get-RemoteMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue).Count

[PSCustomObject]@{
    "On-Premise Mailboxes" = if ($OnPremCount) { $OnPremCount } else { 0 }
    "Remote Mailboxes"    = if ($RemoteCount) { $RemoteCount } else { 0 }
} | Format-Table -AutoSize

Write-Host "`n=== Receive Connectors Using Explicit TLS Certificates ===" -ForegroundColor Cyan
Get-ReceiveConnector | 
    Where-Object { -not [string]::IsNullOrEmpty($_.TlsCertificateName) } | 
    Select-Object Identity, Enabled, FQDN, TlsCertificateName | 
    Format-Table -AutoSize

Write-Host "`n=== Send Connectors Using Explicit TLS Certificates ===" -ForegroundColor Cyan
Get-SendConnector | 
    Where-Object { -not [string]::IsNullOrEmpty($_.TlsCertificateName) } | 
    Select-Object Identity, Enabled, FQDN, TlsCertificateName | 
    Format-Table -AutoSize

Write-Host "`n=== Default Internal SMTP Certificates (Active Directory) ===" -ForegroundColor Cyan
$ExchangeServers = Get-ExchangeServer | Where-Object { $_.IsHubTransportServer -or $_.IsMailboxServer }

foreach ($Server in $ExchangeServers) {
    Write-Host "Server: $($Server.Name)" -ForegroundColor DarkCyan
    
    try {
        $AdObject = Get-ADObject -Identity $Server.DistinguishedName -Properties msExchServerInternalTLSCert -ErrorAction Stop
        if ($AdObject.msExchServerInternalTLSCert) {
            $Cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($AdObject.msExchServerInternalTLSCert)
            
            [PSCustomObject]@{
                Server      = $Server.Name
                Thumbprint  = $Cert.Thumbprint
                FriendlyName= $Cert.FriendlyName
                Subject     = $Cert.Subject
                Expires     = $Cert.NotAfter
            } | Format-Table -AutoSize
        } else {
            Write-Warning "No internal TLS certificate attribute found in Active Directory for $($Server.Name)."
        }
    } catch {
        Write-Error "Failed to retrieve or parse the internal Active Directory TLS certificate for $($Server.Name): $_"
    }
}
