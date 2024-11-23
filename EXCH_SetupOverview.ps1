<#

.DESCRIPTION
This script will give you an overview of how the Exchange is setup, and what configurations have been made. 

.OUTPUTS
ExchangeReport.txt contains all information
DomainChecker.txt contains all information about public DNS entries for each accepted domain

                                            
                                            Domain Controller Information
                                            Exchange Server Information
                                            Mailbox Database configuration
                                            Database Backup Timestamps
                                            Number and type of mailboxes
                                            Transport Rules
                                            AcceptedDomains
                                            Retention Policy
                                            Send Connectors
                                            Receive Connectors
                                            Exchange Certificates
                                            Organization Configuration
                                            Virtual Directory - Urls & Auth
                                            MX, SPF, Random SPF, DMARC, Selector1 and Selector 2 lookups
                                            
                                            
                                            
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
Domain Controller information
###########################################################################

" -ForegroundColor Yellow

$DomainControllers = Get-ADDomainController -filter * | Select Hostname
$Data = @()

foreach ($DomainController in $Domaincontrollers) {

$MyObject = New-Object PSObject -Property @{
Domain = (Get-ADDomain).DNSRoot
Servername = $DomainController.Hostname
ForestLevel = (Get-ADForest).ForestMode
DomainLevel = (Get-ADDomain).DomainMode
OS = (Get-CimInstance -ComputerName $DomainController.Hostname -ClassName Win32_OperatingSystem).Caption
}
$Data += $MyObject
}
$Data | FL Domain, Servername, OS, ForestLevel, DomainLevel

$RecycleBin = get-adoptionalfeature "recycle bin feature";
$Forestmode = (get-adforest).forestmode;
if (($RecycleBin.EnabledScopes).count -eq 0) 
{	
Write-Host "AD Recycle Bin: DISABLED"
}
else
{
Write-Host "AD Recycle Bin: ENABLED"
}

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
Group Policies that might effect Outlook behaviour
###########################################################################

" -ForegroundColor Yellow

$DC = (Get-ADDomainController | Select Name -First 1).Name
Invoke-Command -ComputerName $DC {

$AllGPO = Get-GPO -All -Domain $env:SERDNSDOMAIN
[string[]] $MatchedGPOList = @()

ForEach ($GPO in $AllGPO) { 
    $Report = Get-GPOReport -Guid $GPO.Id -ReportType XML 
    if ($Report -match 'Outlook') { 
        Write-Host "$($GPO.DisplayName)" -ForeGroundColor "Green"
        $MatchedGPOList += "$($GPO.DisplayName)";
} 
  }
    }

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


$ErrorActionPreference = 'SilentlyContinue'
$Domains = Get-AcceptedDomain | Where-Object {$_.DomainName -notlike "*.local" -and $_.DomainName -notlike "*.onmicrosoft.com"}  | Select DomainName
$Result = foreach ($Domain in $Domains) {
$DomainName = ($Domain.DomainName).Address

# MX Records
$MX = nslookup -q=mx $DomainName 8.8.8.8 2>$null
$MXGrabber = $MX | Select-String "mail exchanger"

# SPF Records
$SPF = nslookup -q=txt $DomainName 8.8.8.8 2>$null
$SPFGrabber = $SPF | Select-String "spf"

# DMARC Records
$DMARC = nslookup -q=txt _dmarc.$DomainName 8.8.8.8 2>$null
$DMARCGrabber = $DMARC | Select-String "v=DMARC1"

# Random SPF Record
$RandomSPF = nslookup -q=txt randomspfrecord.$DomainName 8.8.8.8 2>$null
$RandomSPFGrabber = $RandomSPF | Select-String "v=spf1"

# Selector1 CNAME Record
$Selector1 = nslookup -q=cname "selector1._domainkey.$DomainName" 8.8.8.8 2>$null
$Selector1Grabber = $Selector1 | Select-String "canonical name"

# Selector2 CNAME Record
$Selector2 = nslookup -q=cname "selector2._domainkey.$DomainName" 8.8.8.8 2>$null
$Selector2Grabber = $Selector2 | Select-String "canonical name"

# Display Results
Echo "--- DNS Information for $DomainName ---" >> $home\Desktop\DomainChecker.txt

# MX Record
If ($MXGrabber) {
    $MXRecord = $MXGrabber.Line -replace ".*mail exchanger = ", ""
    Echo "MX-Record:
$MXRecord
    " >> $home\Desktop\DomainChecker.txt
} Else {
    Echo "MX-record:
No valid MX-record found for $DomainName
    " >> $home\Desktop\DomainChecker.txt
}

# SPF Record
If ($SPFGrabber) {
    $SPFRecord = $SPFGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
    Echo "SPF-record:
$SPFRecord
    " >> $home\Desktop\DomainChecker.txt
    
} Else {
    Echo "SPF-record:
No valid SPF-record found for $DomainName
    " >> $home\Desktop\DomainChecker.txt
}

# Random SPF Record
If ($RandomSPFGrabber) {
    $RandomSPF = $RandomSPFGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
    Echo "Random SPF:
$RandomSPF
    " >> $home\Desktop\DomainChecker.txt
    
} Else {
    Echo "Random SPF:
No Random SPF found for $DomainName
    " >> $home\Desktop\DomainChecker.txt
}


# DMARC Record
If ($DMARCGrabber) {
    $DMARCRecord = $DMARCGrabber.Line -replace ".*text = ", "" -replace '"', ""  # Clean up result
 Echo "DMARC-record:
$DMARCRecord
    " >> $home\Desktop\DomainChecker.txt
    
} Else {
    Echo "DMARC-record:
No valid DMARC-record found for $DomainName
    " >> $home\Desktop\DomainChecker.txt
}

# Selector1 Record
If ($Selector1Grabber) {
$S1 = $Selector1Grabber.Line -replace ".*canonical name = ", ""
 Echo "Selector1:
$S1
    " >> $home\Desktop\DomainChecker.txt
    
} Else {
    Echo "Selector1:
No CNAME found for Selector1 $DomainName
    " >> $home\Desktop\DomainChecker.txt
}


# Selector2 Rcord
If ($Selector2Grabber) {
$S2 = $Selector2Grabber.Line -replace ".*canonical name = ", ""
 Echo "Selector2:
$S2
    " >> $home\Desktop\DomainChecker.txt
    
} Else {
    Echo "Selector2:
No CNAME found for Selector2 $DomainName
    
    
    " >> $home\Desktop\DomainChecker.txt
}
    }

Write-host "Script completed. 
Find your output files on your desktop here: $home\Desktop\ExchangeReport.txt and $home\Desktop\DomainChecker.txt" -ForegroundColor Green
