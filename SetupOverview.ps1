<#

.DESCRIPTION
This script will give you an overview of how the Exchange is setup, and what configurations have been made. 
There will be two output files "ExchangeReport.txt" & "Mailbox.csv". Both files will be placed on your desktop.

.OUTPUTS
ExchangeReport.txt contains all Exchange configuration
Mailbox.csv contains all the relevant information about mailboxes.

#>


Import-Module ActiveDirectory
Add-PSSnapin *EXC*
Start-Transcript -path $home\Desktop\ExchangeReport.txt -append | out-null

Write-Host "

###########################################################################
Exchange Server Information
###########################################################################

" -ForegroundColor Green

$Serverlist = Get-ExchangeServer | Select Name
$Data = @()

foreach ($Server in $Serverlist) {
$Items = New-Object PSObject -Property @{
Domain = (Get-ADDomain).DNSRoot
Servername = $Server.Name
OS = (Get-CimInstance -ComputerName $Server.Name -ClassName Win32_OperatingSystem).Caption
RAM = (Invoke-command $Server.Name {(systeminfo | Select-String 'Total Physical Memory:').ToString().Split(':')[1].Trim()})
Exchver = (Invoke-command $Server.Name {Get-WmiObject -Class Win32_Product | Where {$_.Caption -eq "Microsoft Exchange Server"}}).Version
InitialPage = (Invoke-command $Server.Name {Get-CimInstance Win32_PageFileSetting}).InitialSize
MaxPage = (Invoke-command $Server.Name {Get-CimInstance Win32_PageFileSetting}).MaximumSize
CFreeGB = (Invoke-Command $Server.Name {(Get-WmiObject win32_logicaldisk -Filter "DeviceID='C:'").FreeSpace / 1gb -as [int]})
Boot = (Invoke-command $Server.Name {systeminfo | find "System Boot Time:"})

}
$Data += $Items
}
$Data | FL Domain, Servername, OS, RAM, Exchver, InitialPage, MaxPage, CFreeGB, Boot



Write-Host "

###########################################################################
MailBox Databases
###########################################################################

" -ForegroundColor Green

Get-MailboxDatabase -Status | fl Name, DatabaseSize, Server, EDBFilePath, LogFolderPath, MasterServerOrAvailabilityGroup


Write-Host "

###########################################################################
Database backup timestamps
###########################################################################

" -ForegroundColor Green

Get-MailboxDatabase -Status | Fl Name, LastFullBackup, LastIncrementalBackup


Write-Host "

###########################################################################
Number of mailboxes
###########################################################################

" -ForegroundColor Green

Write-Host "UserMailboxes"
(Get-Mailbox -RecipientTypeDetails UserMailbox -Resultsize Unlimited).count

Write-Host "SharedMailboxes"
(Get-Mailbox -RecipientTypeDetails SharedMailbox -Resultsize Unlimited).count

Write-Host "RoomMailboxes"
(Get-Mailbox -RecipientTypeDetails RoomMailbox -Resultsize Unlimited).count

Write-Host "PublicFolders"
(Get-PublicFolder "\" -Recurse -ErrorAction SilentlyContinue).count

Write-Host "

###########################################################################
TransportRules
###########################################################################

" -ForegroundColor Green

Get-TransportRule -WarningAction SilentlyContinue | fl Name, State



Write-Host "

###########################################################################
AcceptedDomains
###########################################################################

" -ForegroundColor Green

Get-AcceptedDomain | fl Name

Write-Host "

###########################################################################
RetentionPolicy
###########################################################################

" -ForegroundColor Green

Get-Retentionpolicy | fl Name


Write-Host "

###########################################################################
Send Connectors
###########################################################################

" -ForegroundColor Green

Get-SendConnector | fl name, Smarthosts


Write-Host "

###########################################################################
Receive Connectors
###########################################################################

" -ForegroundColor Green


Get-ReceiveConnector | fl Name, Enabled, RemoteIPRanges



Write-Host "

###########################################################################
Exchange Certificates
###########################################################################

" -ForegroundColor Green


Get-ExchangeCertificate | fl Thumbprint, IsSelfSigned, Subject, Services, Notafter, NotBefore


Write-Host "

###########################################################################
Exchange URL's
###########################################################################


" -ForegroundColor Green

Write-Host "Autodiscover"
Get-ClientAccessServer -WarningAction SilentlyContinue -Identity "$env:COMPUTERNAME" | fl AutodiscoverServiceInternalURI

Write-Host "OWA (Outlook Web Application)"
Get-OwaVirtualDirectory -Identity "$env:COMPUTERNAME\OWA (Default Web Site)" | fl InternalURL, ExternalURL

Write-Host "ECP (Exchange Control Panel)"
Get-ECPVirtualDirectory -Identity "$env:COMPUTERNAME\ECP (Default Web Site)" | fl InternalURL, ExternalURL

Write-Host "EWS (Exchange Web Services)"
Get-WebServicesVirtualDirectory -Identity "$env:COMPUTERNAME\EWS (Default Web Site)" | fl InternalURL, ExternalURL

Write-Host "MAPI"
Get-MapiVirtualDirectory -Identity "$env:COMPUTERNAME\MAPI (Default Web Site)" | fl InternalURL, ExternalURL

Write-Host "OAB (Offline Address Book)"
Get-OABVirtualDirectory -Identity "$env:COMPUTERNAME\OAB (Default Web Site)" | fl InternalURL, ExternalURL

Write-Host "EAS (Exchange Active Sync)"
Get-ActiveSyncVirtualDirectory -Identity "$env:COMPUTERNAME\Microsoft-Server-ActiveSync (Default web site)" | fl InternalURL, ExternalURL

Write-Host "Outlook Anywhere"
Get-OutlookAnywhere -Identity "$env:COMPUTERNAME\rpc (Default web site)" | Fl InternalHostname, ExternalHostname


Write-Host "

###########################################################################
Exchange Authentication protocols
###########################################################################


" -ForegroundColor Green

Write-Host "Autodiscover"
Get-ExchangeServer $env:computername | Get-AutodiscoverVirtualDirectory | fl InternalAuthenticationMethods, ExternalAuthenticationMethods

Write-Host "OWA (Outlook Web Application"
Get-ExchangeServer $env:computername | Get-OWAVirtualDirectory | fl InternalAuthenticationMethods, ExternalAuthenticationMethods

Write-Host "ECP (Exchange Control Panel)"
Get-ExchangeServer $env:computername | Get-ECPVirtualDirectory | fl InternalAuthenticationMethods, ExternalAuthenticationMethods

Write-Host "EWS (Exchange Web Services)"
Get-ExchangeServer $env:computername | Get-WebServicesVirtualDirectory | fl InternalAuthenticationMethods, ExternalAuthenticationMethods

Write-Host "MAPI"
Get-ExchangeServer $env:computername | Get-MapiVirtualDirectory | fl InternalAuthenticationMethods, ExternalAuthenticationMethods

Write-Host "Outlook Anywhere"
Get-ExchangeServer $env:computername | Get-OutlookAnywhere | fl InternalClientAuthenticationMethod, ExternalClientAuthenticationMethod, IISAuthenticationMethods

Stop-Transcript | out-null


Write-Host "

###########################################################################
Mailbox data export
###########################################################################

" -ForegroundColor Green


$UserList = Get-MailBox -Resultsize Unlimited
$ExportList = @()

foreach ($User in $UserList) {

$Collection = New-Object PSObject -Property @{

FullAccess = (Get-MailboxPermission $User | Where {$_.AccessRights -EQ "FullAccess" -and -not ($_.User -like “NT AUTHORITY\*”)}).User
SendAs = (Get-Mailbox $User | Get-ADPermission | Where {$_.ExtendedRights -like "Send-As" -and -not ($_.User -like “NT AUTHORITY\*”)}).User
Name = (Get-MailboxStatistics $User).DisplayName
Size = (Get-MailboxStatistics $User).TotalItemSize.Value.ToMB()
Deleted = (Get-MailboxStatistics $User).TotalDeletedItemSize.Value.ToMB()
Total = $null
Username = (Get-Mailbox $User).SamAccountName
Email = (Get-Mailbox $User).PrimarySmtpAddress
Type = (Get-Mailbox $User).RecipientTypeDetails
DBName = (Get-MailboxStatistics $User).DatabaseName
SMTPAlias = (Get-Mailbox $User).Emailaddresses
LastLogon = (Get-MailboxStatistics $User).LastLogonTime
ADEnabled = (Get-ADUser $User.SamAccountName).Enabled

}

$ExportList += $Collection

}

# Select fields in specific order rather than random.
$ExportList | Select Username, Name, Email, Type, Size, Deleted, Total, DBName, LastLogon, ADEnabled, {$_.FullAccess}, {$_.SendAs}, {$_.SMTPAlias} | 
Export-csv $home\Desktop\ExchangeExport.csv -NoTypeInformation -Encoding Unicode
