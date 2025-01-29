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
Exchange Version: $ExchVer
Forest Level: $ForestLevel
Domain Level: $DomainLevel
Mailboxes ONPREM: $Mailbox
Mailboxes EXO: $RemoteMailbox
Hybrid Enabled: $Hybrid


"

# Send mail
Send-MailMessage -From $Sender -to exchangeteam@itm8.com -Subject $Subject -SmtpServer localhost -Body $Body
