# Parameters
Add-PSSnapin *EXC*
Import-Module ActiveDirectory

$Domain = (Get-Accepteddomain | Where {$_.Default -EQ "True"}).Name
$Sender = "ITM8_Exchange@$domain"
$Subject = "Customer: $Domain - Server: $Env:computername"
$Mailbox = (Get-Mailbox).Count
$RemoteMailbox = (Get-RemoteMailbox).count
$ExchVer = (Get-Command Exsetup.exe).Version
$ForestLevel = (Get-ADForest).ForestMode
$DomainLevel = (Get-ADDomain).DomainMode
$HybridConfig = (Get-HybridConfiguration).WhenChanged
$DisplayExchange = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "Microsoft Exchange Server 20*"}).Displayname
$RConnector = (Get-ReceiveConnector | Where {$_.TlsCertificateName -NE $null}).Name -join ", "
$SConnector = (Get-SendConnector | Where {$_.TlsCertificateName -NE $null}).Name -join ", "
$Certificates = (Get-ExchangeCertificate | Where {$_.Services -NE "none"}).FriendlyName -join ", "

$Hybrid = Get-HybridConfiguration
If ($Hybrid)
{
$Hybrid = "True"
}
Else
{
$Hybrid = "False"
}

# Output
$Body = "
Servername: $env:computername
Exchange: $DisplayExchange
Version: $ExchVer

Forest Level: $ForestLevel
Domain Level: $DomainLevel

Mailboxes ONPREM: $Mailbox
Mailboxes EXO: $RemoteMailbox

Hybrid Enabled: $Hybrid
Hybrid Changed: $HybridConfig

Receive Connectors with cert:
$RConnector

Send Connectors with cert:
$SConnector

Certificates with services:
$Certificates

"
# Send mail
Send-MailMessage -From $Sender -to exchangeteam@itm8.com -Subject $Subject -SmtpServer localhost -Body $Body
