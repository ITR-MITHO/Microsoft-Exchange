$Domain = (Get-Accepteddomain | Where {$_.Default -EQ "True"}).Name
$Sender = "ITM8_Exchange@$domain"
$Subject = "Customer: $Domain - Server: $Env:computername"

$ExchangeDisplay = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall | ForEach-Object {
    $DisplayName = $_.GetValue("DisplayName")
    if ($DisplayName -match "Microsoft Exchange Server 20*") {
        [PSCustomObject]@{
            DisplayName = $DisplayName}
                                        }
                                            }

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


$Body = "
Servername: $env:computername
Exchange Display: $ExchangeDisplay
Exchange Version: $ExchVer
Forest Level: $ForestLevel
Domain Level: $DomainLevel
Mailboxes ONPREM: $Mailbox
Mailboxes EXO: $RemoteMailbox
Hybrid Enabled: $Hybrid


"
Send-MailMessage -From $Sender -to exchangeteam@itm8.com -Subject $Subject -SmtpServer localhost -Body $Body




