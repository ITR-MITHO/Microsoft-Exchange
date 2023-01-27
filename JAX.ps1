<#
J.A.X Version 2.0


.DESCRIPTION
The script is designed to be a multi-tool for Microsoft Exchange

.NOTES
It includes the following features; 

CU version Check
Show Exchange URL's
Change Exchange URL's
HealthCheck (Will prompt for e-mail to test external mailflow)
Mailbox Export Report
Status on backup
Exchange Server Setup Report

To use the script, type one of the numbers corresponding to the tool; 1, 2, 3, 4, 5, 6, 7 or 0 for exit.
#>

Function Start-JAX {
    CLS    
    $Function = Read-Host "
    
    Enter"1" - CU Version Check
    Enter "2" - Show Exchange URL's
    Enter "3" - Change Exchange URL's
    Enter "4" - HealthCheck
    Enter "5" - Mailbox Export Report
    Enter "6" - Backup Status
    Enter "7" - Exchange Server Setup Report
    Enter "0" - Exit
    "

    If ($Function -EQ "1") 
    {

<# Arrays #>
Write-Host "Starting CU check" -ForegroundColor Yellow
Add-PSSnapin *EXC*
$Display = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "Microsoft Exchange Server 20*"}
$Exchange2019 = "Microsoft Exchange Server 2019 Cumulative Update 11"
$Exchange2016 = "Microsoft Exchange Server 2016 Cumulative Update 23"
$Exchange2013 = "Microsoft Exchange Server 2013 Cumulative Update 23"
$Exchange2010 = "Microsoft Exchange Server 2010"
$Exchange2007 = "Microsoft Exchange Server 2007"
$Servers = Get-ExchangeServer -Identity $Env:ComputerName
$Output = $Display.DisplayName

<# Exchange Server 2019 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*Microsoft Exchange Server 2019*")
{

if ($Display.DisplayName -GE "$Exchange2019")
{

Write-Host "$Output - SUPPORTED" -ForegroundColor Green

}
Else
{

Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}

   }
        }
    

<# Exchange Server 2016 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*Microsoft Exchange Server 2016*")
{

if ($Display.DisplayName -GE "$Exchange2016")
{

Write-Host "$Output - SUPPORTED" -ForegroundColor Green

}
Else
{

Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}

   }
        }


<# Exchange Server 2013 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*Microsoft Exchange Server 2013*")
{

if ($Display.DisplayName -GE "$Exchange2013")
{

Write-Host "$Output - SUPPORTED" -ForegroundColor Green

}
Else
{

Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}

   }
        }


<# Exchange Server 2010 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*$Exchange2010*")

{

        Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}
    
        }


<# Exchange Server 2007 #>
Foreach ($S in $Servers)
        {

If ($Display -like "*$Exchange2007*")

{

        Write-host "$Output - NOT SUPPORTED" -ForegroundColor Red

}
    
        }
        Write-Host "Type"Start-Jax" to get back to the menu." -ForegroundColor Yellow
    }



    If ($Function -EQ "2") 
    {

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


        Write-Host "Type"Start-Jax" to get back to the menu." -ForegroundColor Yellow
    }

    

    If ($Function -EQ "3") 
    {

        $DomainName = Read-Host "Enter domain name here. (e.g. domain.com)"
        $URL = "https://mail.$DomainName"
                                                                                                   
        Set-ClientAccessServer -Identity "$env:COMPUTERNAME" -AutodiscoverServiceInternalURI "https://autodiscover.$DomainName/Autodiscover/Autodiscover.xml"
        Set-OwaVirtualDirectory -Identity "$env:COMPUTERNAME\owa (Default Web Site)" -InternalUrl "$URL/owa" -ExternalUrl "$URL/owa"
        Set-MapiVirtualDirectory -Identity "$env:COMPUTERNAME\mapi (Default Web Site)" -InternalUrl "$URL/mapi" -ExternalUrl "$URL/mapi"
        Set-OabVirtualDirectory -Identity "$env:COMPUTERNAME\oab (Default Web Site)" -InternalUrl "$URL/oab" -ExternalUrl "$URL/oab"
        Set-EcpVirtualDirectory -Identity "$env:COMPUTERNAME\ecp (Default Web Site)" -InternalUrl "$URL/ecp" -ExternalUrl "$URL/ecp"
        Set-WebServicesVirtualDirectory -Identity "$env:COMPUTERNAME\ews (Default Web Site)" -InternalUrl "$URL/EWS/Exchange.asmx" -ExternalUrl "$URL/EWS/Exchange.asmx"
        Set-ActiveSyncVirtualDirectory -Identity "$env:COMPUTERNAME\Microsoft-Server-ActiveSync (Default web site)" -InternalUrl "$URL/Microsoft-Server-ActiveSync" -ExternalUrl "$URL/Microsoft-Server-ActiveSync"
        Get-OutlookAnywhere -Server "$env:COMPUTERNAME" | Set-OutlookAnywhere -ExternalHostname "mail.$DomainName" -InternalHostname "mail.$DomainName" -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $true -DefaultAuthenticationMethod NTLM
        
        Write-Host "Type"Start-Jax" to get back to the menu." -ForegroundColor Yellow
    }



    If ($Function -EQ "4") 
    {

#>

$Recipient = Read-Host "Enter EXTERNAL e-mailaddress (e.g. email@gmail.com)"
$mydomain = (Get-ADDomain).DNSRoot
Send-MailMessage -To $Recipient -From ExchangeHealth@$mydomain -Subject "Testing mailflow" -Body "
If you have rececived this e-mail, the mail-flow from your on-premises Exchange is working.
Best regards, 
$env:computername.$mydomain
" -SmtpServer "$env:computername.$mydomain"


# MessageQueue
$Queue = (Get-ExchangeServer | Get-Message -ErrorAction SilentlyContinue).count
$Date = Get-Date -Format "dd-MM-yy HH:mm"
If ($Queue -GT 100)
{
Write-Host "Over 100 e-mails are in queue!" -ForegroundColor Red
}
Else
{
Write-Host "Exchange Message queue is healthy.
" -ForegroundColor Green
}


# ComponentState
$Component = Get-ServerComponentState -Identity $env:computername | Where {$_.Component -NE "ForwardSyncDaemon" -and $_.Component -NE "ProvisioningRps"}
if ($Component | Where {$_.State -eq "inactive"})
{
Write-Host "Exchange componenets are inactive!" -ForegroundColor Red
$Component
}
Else
{
Write-Host "Exchange Server Components is running
" -ForegroundColor Green
}


# ServiceHealth
$ServiceHealth = Test-ServiceHealth $env:computername
if ($ServiceHealth | Where {$_.RequiredServicesRunning -NE $true})
{
Write-Host "Microsoft Exchange Services are not running!" -ForegroundColor Red
$ServiceHealth
}
else
{
Write-Host "Microsoft Exchange Services are running
" -ForegroundColor Green
}


# MapiConnectivity
$MAPIConnectivity = Test-MAPIConnectivity
If ($MAPIConnectivity | Where {$_.Result -EQ "Failed"})
{
Write-Host "MapiConnectivity failed." -ForegroundColor Red
$MapiConnectivity
}
Else
{
Write-Host "MAPIConnectivityTest passed
" -ForegroundColor Green
}


# DAGReplicationHealth
$DAGTest = Test-ReplicationHealth $env:computername | Where {$_.Result -like "*failed*"} | Select Server, Check, Result
$DAG = Get-DatabaseAvailabilityGroup
If ($DAG -ne $null)
{

Write-Host "Exchange DAG Found.. Testing replication" -ForegroundColor Yellow
Test-ReplicationHealth $env:computername | Select Server, Check, Result
sleep 5
}
Else
{
Write-Host "No Exchange DAG found, skipping replication check.
" -ForegroundColor Yellow
}


If ($DAG -ne $null)
{

if ($DAGTest -ne $null)
{
    Write-Host "Exchange DAG replication is unhealthy! (Test-replicationhealth)
    
    " -ForegroundColor Red

}

}

Write-Host "Type"Start-Jax" to get back to the menu." -ForegroundColor Yellow

    }

    If ($Function -EQ "5") 
    {
Write-Host "Starting Mailbox Export Report" -ForegroundColor Yellow
        Add-PSSnapin *EXC*
        Import-Module ActiveDirectory
        
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
        Export-csv C:\Users\$env:username\Desktop\ExchangeExport.csv -NoTypeInformation -Encoding Unicode


        Write-Host "You can find the .csv-file here: $home\desktop\ExchangeExport.csv" -ForegroundColor Yellow
        Write-Host "Type"Start-Jax" to get back to the menu." -ForegroundColor Yellow
    }

    If ($Function -EQ "6")
    {

<# Variables #>
Add-PSSnapin *EXC* 
$Timestamp = Get-MailboxDatabase -Status
$Date = (Get-date).AddDays(-1) # Veeam Full Snapshot Backup
$Date3 = (Get-date).AddDays(-3) # Incremental & Differential
$Date8 = (Get-date).AddDays(-8) # Servers with incremental and differential
 
<# Veeam Full Snapshot Backup #>

If (-not $Timestamp.LastDifferentialBackup -and -not $Timestamp.LastIncrementalBackup)
    {
    If ($Timestamp.LastFullBackup -LT $Date)
{
    Write-Host "Veeam full backup is over 24-hours old!" -ForegroundColor Red
}
else 
{
    Write-Host "Veaam full OK"
}
    }
 
<# Incremental Backup #>
If (-not $Timestamp.LastDifferentialBackup)
{
    If (-not $Timestamp.LastIncrementalBackup)
    {

    }
    else    
{
        
If ($Timestamp.LastIncrementalBackup -LT $Date3)
 
{
Write-Host "Incremental backup is over 3 days old!" -ForegroundColor Red
}
        }
            }
 
<# Differential Backup #>
If (-not $Timestamp.LastIncrementalBackup)
 
{
    If (-not $Timestamp.LastDifferentialBackup)
    {

    }
 
Else 
   
{
 
If ($Timestamp.LastDifferentialBackup -LT $Date3)
    
{
Write-Host "Differential backup is over 3 days old!" -ForegroundColor Red
}
        }
            }
 
<# 8-day Fullbackup #>
 
If ($Timestamp.LastFullBackup -LT $Date8)
{
Write-Host "Full backup is over 8 days old!" -ForegroundColor Red        
}
else 
{    
Write-Host "8-day full backup is OK"   
}

Write-Host "Type"Start-Jax" to get back to the menu." -ForegroundColor Yellow
    } 



    If ($Function -EQ "7") 
    {

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
        (Get-Mailbox -RecipientTypeDetails UserMailbox).count
        
        Write-Host "SharedMailboxes"
        (Get-Mailbox -RecipientTypeDetails SharedMailbox).count
        
        Write-Host "RoomMailboxes"
        (Get-Mailbox -RecipientTypeDetails RoomMailbox).count
        
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

        Write-Host "The report can be found here: $home\Desktop\ExchangeReport.txt" -ForegroundColor Yellow
        Write-Host "Type"Start-Jax" to get back to the menu." -ForegroundColor Yellow
    }


    If ($Function -EQ "0")
    {

        exit

    } 

}
Start-JAX
