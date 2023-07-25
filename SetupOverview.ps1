<#

.DESCRIPTION
This script will give you an overview of how the Exchange is setup, and what configurations have been made. 

.OUTPUTS
ExchangeReport.txt contains all information

                                            
                                            
                                            Exchange Server Information
                                            Mailbox Database configuration
                                            Database Backup Timestamps
                                            Number and type of mailboxes
                                            Mailflow statistics
                                            Transport Rules
                                            AcceptedDomains
                                            Retention Policy
                                            Send Connectors
                                            Receive Connectors
                                            Exchange Certificates
                                            Organization Configuration
                                            Virtual Directory - Urls & Auth
                                            DMARC records for accepted domains
                                            
                                            
                                            
#>

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
write-host "Script is not running as Administrator" -ForegroundColor Yellow
Break
}


Import-Module ActiveDirectory
Add-PSSnapin *EXC*
Start-Transcript -path $home\Desktop\ExchangeReport.txt -append | out-null

Write-Host "

###########################################################################
Exchange Server Information
###########################################################################

" -ForegroundColor Yellow

$Serverlist = Get-ExchangeServer | Select Name
$Data = @()

foreach ($Server in $Serverlist) {
$Items = New-Object PSObject -Property @{
Domain = (Get-ADDomain).DNSRoot
Servername = $Server.Name
OS = (Get-CimInstance -ComputerName $Server.Name -ClassName Win32_OperatingSystem).Caption
RAM = (Invoke-command $Server.Name {(systeminfo | Select-String 'Total Physical Memory:').ToString().Split(':')[1].Trim()})
Exchver = (Invoke-command $Server.Name {Get-Command Exsetup.exe | ForEach {$_.FileVersionInfo}}).FileVersion
InitialPage = (Invoke-command $Server.Name {Get-CimInstance Win32_PageFileSetting}).InitialSize
MaxPage = (Invoke-command $Server.Name {Get-CimInstance Win32_PageFileSetting}).MaximumSize
}
$Data += $Items
}
$Data | FL Domain, Servername, OS, RAM, Exchver, InitialPage, MaxPage

Write-Host "

###########################################################################
Mailbox Databases
###########################################################################

" -ForegroundColor Yellow

Get-MailboxDatabase -Status | fl Name, DatabaseSize, Server, EDBFilePath, LogFolderPath, MasterServerOrAvailabilityGroup, DeletedItemRetention, MailboxRetention, CircularLoggingEnabled, AvailableNewMailboxSpace


Write-Host "

###########################################################################
Database backup timestamps
###########################################################################

" -ForegroundColor Yellow

Get-MailboxDatabase -Status | Fl Name, LastFullBackup, LastIncrementalBackup, LastDifferentialBackup


Write-Host "

###########################################################################
Number of mailboxes
###########################################################################

" -ForegroundColor Yellow

Write-Host "UserMailboxes"
(Get-Mailbox -RecipientTypeDetails UserMailbox -Resultsize Unlimited).count

Write-Host "SharedMailboxes"
(Get-Mailbox -RecipientTypeDetails SharedMailbox -Resultsize Unlimited).count

Write-Host "RoomMailboxes"
(Get-Mailbox -RecipientTypeDetails RoomMailbox -Resultsize Unlimited).count

Write-Host "PublicFolders"
(Get-PublicFolder "\" -Recurse -ErrorAction SilentlyContinue).count

Write-Host "DynamicDistributionGroups (DDG does not work in hybrid mode)"
(Get-DynamicDistributionGroup -Resultsize Unlimited).Count




Write-Host "

###########################################################################
Mailflow last 24 hours
###########################################################################

" -ForegroundColor Yellow

$24Hours = (Get-Date).AddDays(-1)
$Trace = Get-ExchangeServer | Get-MessageTrackingLog -Start $24Hours -ResultSize Unlimited
$Send = ($Trace | Where {$_.EventID -EQ "Send"}).count
Write-Host "$Send e-mails sent within the last 24 hours. "

$Deliver = ($Trace | Where {$_.EventID -EQ "Deliver"}).count
Write-Host "$Deliver e-mails delivered within the last 24 hours"



Write-Host "

###########################################################################
TransportRules
###########################################################################

" -ForegroundColor Yellow

Get-TransportRule -WarningAction SilentlyContinue | fl Name, State



Write-Host "

###########################################################################
AcceptedDomains
###########################################################################

" -ForegroundColor Yellow

Get-AcceptedDomain | fl Name, DomainType, DomainName

Write-Host "

###########################################################################
RetentionPolicy
###########################################################################

" -ForegroundColor Yellow
$Retention = Get-Retentionpolicy | Select Name, RetentionPolicyTagLinks

Foreach ($R in $Retention)
{

$RetentionName = $R.Name
$RetentionTag = $R.RetentionPolicyTagLinks
$RetentionCount = (Get-Mailbox -ResultSize unlimited | Where {$_.RetentionPolicy -eq "$RetentionName"}).count

Write-Host " 

$RetentionName is assiged to $RetentionCount mailboxes.
PolicyTag: $RetentionTag

" 

}

Write-Host "

###########################################################################
Send Connectors
###########################################################################

" -ForegroundColor Yellow

Get-SendConnector | fl name, Smarthosts, AddressSpaces, Enabled, ProtocolLoggingLevel


Write-Host "

###########################################################################
Receive Connectors
###########################################################################

" -ForegroundColor Yellow


Get-ReceiveConnector | fl Name, Enabled, RemoteIPRanges, ProtocolLoggingLevel, PermissionGroups



Write-Host "

###########################################################################
Exchange Certificates
###########################################################################

" -ForegroundColor Yellow


Get-ExchangeCertificate | fl Services, Thumbprint, IsSelfSigned, Subject, Notafter, NotBefore


Write-Host "

###########################################################################
OrganizationConfig
###########################################################################

" -ForegroundColor Yellow

$Kerb = Get-ClientAccessServer $env:COMPUTERNAME -IncludeAlternateServiceAccountCredentialStatus -WarningAction SilentlyContinue | Select AlternateServiceAccountConfiguration
If ($Kerb.AlternateServiceAccountConfiguration -like "Latest: <n*")
{
Write-Host "KerberosEnabled: False

"
}
Else
{
Write-Host "KerberosEnabled: True

"
}

$Hybrid = Get-HybridConfiguration
If ($Hybrid)
{
Write-Host "HybridEnabled: True"
}
Else
{
Write-Host "HybridEnabled: False"
}

Get-OrganizationConfig | fl OAuth2ClientProfileEnabled, MitigationsEnabled, MapiHttpEnabled


Write-Host "
###########################################################################
DMARC Records for all domains.
###########################################################################
" -ForegroundColor Yellow
$ErrorActionPreference = 'SilentlyContinue'
Sleep 5
$Domains = Get-AcceptedDomain | Where DomainName -notlike "*.local" | Select DomainName
$Result = foreach ($Domain in $Domains) {
$ActualDomain = $Domain.DomainName
    Echo "---------------------- $ActualDomain ----------------------"
    Echo ""
    Echo "DMARC TXT Record:"
    (nslookup -q=txt _dmarc.$ActualDomain | Select-String "DMARC1") -replace "`t", ""
    Echo ""

}
$Result


Write-Host "
###########################################################################
Virtual Directories
###########################################################################
" -ForegroundColor Yellow

Write-Host "Autodiscover"
Get-ClientAccessServer -WarningAction SilentlyContinue -Identity "$env:COMPUTERNAME" | fl AutodiscoverServiceInternalURI
Get-ExchangeServer $env:computername | Get-AutodiscoverVirtualDirectory | fl InternalAuthenticationMethods, ExternalAuthenticationMethods

Write-Host "OWA (Outlook Web Application)"
Get-OwaVirtualDirectory -Identity "$env:COMPUTERNAME\OWA (Default Web Site)" | fl InternalURL, ExternalURL, InternalAuthenticationMethods, ExternalAuthenticationMethods 

Write-Host "ECP (Exchange Control Panel)"
Get-ECPVirtualDirectory -Identity "$env:COMPUTERNAME\ECP (Default Web Site)" | fl InternalURL, ExternalURL, InternalAuthenticationMethods, ExternalAuthenticationMethods 

Write-Host "EWS (Exchange Web Services)"
Get-WebServicesVirtualDirectory -Identity "$env:COMPUTERNAME\EWS (Default Web Site)" | fl InternalURL, ExternalURL, InternalAuthenticationMethods, ExternalAuthenticationMethods 

Write-Host "MAPI"
Get-MapiVirtualDirectory -Identity "$env:COMPUTERNAME\MAPI (Default Web Site)" | fl InternalURL, ExternalURL, InternalAuthenticationMethods, ExternalAuthenticationMethods 

Write-Host "OAB (Offline Address Book)"
Get-OABVirtualDirectory -Identity "$env:COMPUTERNAME\OAB (Default Web Site)" | fl InternalURL, ExternalURL, InternalAuthenticationMethods, ExternalAuthenticationMethods 

Write-Host "EAS (Exchange Active Sync)"
Get-ActiveSyncVirtualDirectory -Identity "$env:COMPUTERNAME\Microsoft-Server-ActiveSync (Default web site)" | fl InternalURL, ExternalURL

Write-Host "Outlook Anywhere"
Get-OutlookAnywhere -Identity "$env:COMPUTERNAME\rpc (Default web site)" | Fl InternalHostname, ExternalHostname
Get-ExchangeServer $env:computername | Get-OutlookAnywhere | fl InternalClientAuthenticationMethod, ExternalClientAuthenticationMethod, IISAuthenticationMethods

Stop-Transcript | out-null

Write-host "Script completed. 
Find your output file on your desktop here: $home\Desktop\ExchangeReport.txt" -ForegroundColor Green
