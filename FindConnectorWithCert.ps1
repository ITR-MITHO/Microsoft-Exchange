<#  

The script will check if a certificate is used directly on connectors and if Hybrid is enabled on a server where mailboxes are not present. 

#>

Add-PSSnapin *EXC*
# Mailboxes on-prem
$Mailbox = (Get-Mailbox).count
If ($Mailbox)
{
    Write-Host "Found $Mailbox mailboxes on-prem" -ForegroundColor yellow
}


# Hybrid Configuration
$Hybrid = Get-HybridConfiguration
If ($Hybrid)
{
    Write-Host "Hybrid is configured on the server" -ForegroundColor yellow
}
Else
{
    Write-Host "No Hybrid Configuration" -ForegroundColor Green
}

# Receive Connectors
$Receive = Get-ReceiveConnector | Where {$_.TlsCertificateName -NE $null} | Select Identity, TlsCertificateName, ProtocolLoggingLevel
If ($Receive)
{
Echo '
# Receive Connectors #'
$Receive
}
Else
{
    Write-Host "No Receive Connectors with certificates" -ForegroundColor Yellow
}

# Send Connectors
$Send = Get-SendConnector | Where {$_.TlsCertificateName -NE $null} | Select Identity, TlsCertificateName, ProtocolLoggingLevel
If ($Send)
{
Echo '
# Send Connectors # '
$Send
}
Else
{
    Write-Host "No send Connectors with certificates" -ForegroundColor Yellow
}
