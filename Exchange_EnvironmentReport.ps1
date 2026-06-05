<#
.SYNOPSIS
    Optimized Exchange Environment Script.

.DESCRIPTION
    Generates a structured HTML report of the Exchange configuration. 
#>

$ErrorActionPreference = 'SilentlyContinue'
Import-Module ActiveDirectory
Add-PSSnapin *EXC*

$ReportPath = "$env:USERPROFILE\Desktop\ExchangeReport.html"

$CSS = "<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 14px; background-color: #f4f4f4; color: #333; margin: 20px; }
    h1 { color: #005A9E; border-bottom: 2px solid #005A9E; padding-bottom: 5px; }
    h2 { color: #107C41; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 20px; background-color: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
    th { background-color: #005A9E; color: white; }
    tr:nth-child(even) { background-color: #f9f9f9; }
</style>"

$HTMLBody = "<h1>Exchange Environment Documentation</h1><p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>"

# ---------------------------------------------------------
# Domain Controller Information
# ---------------------------------------------------------
Write-Host "Gathering Domain Controller Info..." -ForegroundColor Cyan
$DCData = Get-ADDomainController -Filter * | ForEach-Object {
    [PSCustomObject]@{
        Domain      = (Get-ADDomain).DNSRoot
        Servername  = $_.Hostname
        ForestLevel = (Get-ADForest).ForestMode
        DomainLevel = (Get-ADDomain).DomainMode
        OS          = (Get-CimInstance -ComputerName $_.Hostname -ClassName Win32_OperatingSystem).Caption
    }
}
$RecycleBin = Get-ADOptionalFeature "Recycle Bin Feature"
$HTMLBody += "<h2>Domain Controller Information (AD Recycle Bin: $(if($RecycleBin.EnabledScopes.Count -gt 0){'Enabled'}else{'Disabled'}))</h2>"
$HTMLBody += $DCData | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Exchange Server Information
# ---------------------------------------------------------
Write-Host "Gathering Exchange Server Info..." -ForegroundColor Cyan
$WANIP = (Invoke-RestMethod -Uri "http://ifconfig.me/ip" -UseBasicParsing).Trim()
$ServerData = Get-ExchangeServer | ForEach-Object {
    $compName = $_.Name
    $IPv4 = ([System.Net.Dns]::GetHostAddresses($compName) | Where-Object AddressFamily -eq 'InterNetwork').IPAddressToString -join ', '

    [PSCustomObject]@{
        Servername  = $compName
        IPv4        = $IPv4
        WANIP       = $WANIP
        OS          = (Get-CimInstance -ComputerName $compName -ClassName Win32_OperatingSystem).Caption
        RAM_GB      = [math]::Round((Get-CimInstance -ComputerName $compName -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
        ExchangeVer = (Get-Command Exsetup.exe | ForEach {$_.FileVersionInfo}).FileVersion
    }
}
$HTMLBody += "<h2>Exchange Server Information</h2>"
$HTMLBody += $ServerData | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Organization Configuration
# ---------------------------------------------------------
Write-Host "Gathering Organization Configuration..." -ForegroundColor Cyan
$Kerb = Get-ClientAccessServer $env:COMPUTERNAME -IncludeAlternateServiceAccountCredentialStatus -WarningAction SilentlyContinue
$Hybrid = Get-HybridConfiguration -ErrorAction SilentlyContinue
$OrgConfig = Get-OrganizationConfig

$OrgData = [PSCustomObject]@{
    KerberosEnabled            = if ($Kerb.AlternateServiceAccountConfiguration -like "*Latest: <n*") { $false } else { $true }
    HybridEnabled              = if ($Hybrid) { $true } else { $false }
    OAuth2ClientProfileEnabled = $OrgConfig.OAuth2ClientProfileEnabled
    MitigationsEnabled         = $OrgConfig.MitigationsEnabled
    MapiHttpEnabled            = $OrgConfig.MapiHttpEnabled
}
$HTMLBody += "<h2>Organization Configuration</h2>"
$HTMLBody += $OrgData | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Mailbox Databases & Backups
# ---------------------------------------------------------
Write-Host "Gathering Mailbox Database Info..." -ForegroundColor Cyan
$DBData = Get-MailboxDatabase -Status | Select-Object Name, DatabaseSize, Server, CircularLoggingEnabled, LastFullBackup, LastIncrementalBackup
$HTMLBody += "<h2>Mailbox Databases & Backups</h2>"
$HTMLBody += $DBData | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Mailbox Statistics
# ---------------------------------------------------------
Write-Host "Gathering Mailbox Statistics..." -ForegroundColor Cyan
$AllMailboxes = Get-Mailbox -ResultSize Unlimited
$MailboxCounts = [PSCustomObject]@{
    UserMailboxes          = ($AllMailboxes | Where-Object RecipientTypeDetails -eq 'UserMailbox').Count
    RemoteMailboxes        = (Get-RemoteMailbox).Count
    SharedMailboxes        = ($AllMailboxes | Where-Object RecipientTypeDetails -eq 'SharedMailbox').Count
    RoomMailboxes          = ($AllMailboxes | Where-Object RecipientTypeDetails -eq 'RoomMailbox').Count
    PublicFolders          = (Get-PublicFolder "\" -Recurse).Count
    DynamicDistGroups      = (Get-DynamicDistributionGroup -ResultSize Unlimited).Count
}
$HTMLBody += "<h2>Mailbox Statistics</h2>"
$HTMLBody += $MailboxCounts | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Accepted Domains
# ---------------------------------------------------------
Write-Host "Gathering Accepted Domains..." -ForegroundColor Cyan
$HTMLBody += "<h2>Accepted Domains</h2>"
$HTMLBody += Get-AcceptedDomain | Select-Object Name, DomainType, DomainName | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Retention Policies
# ---------------------------------------------------------
Write-Host "Gathering Retention Policies..." -ForegroundColor Cyan
$RetData = Get-RetentionPolicy | ForEach-Object {
    $policyName = $_.Name
    [PSCustomObject]@{
        PolicyName     = $policyName
        AssignedCount  = ($AllMailboxes | Where-Object RetentionPolicy -eq $policyName).Count
        PolicyTagLinks = $_.RetentionPolicyTagLinks -join ', '
    }
}
$HTMLBody += "<h2>Retention Policies</h2>"
$HTMLBody += $RetData | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Transport, Connectors & Certificates
# ---------------------------------------------------------
Write-Host "Gathering Transport & Certificates..." -ForegroundColor Cyan

$HTMLBody += "<h2>Transport Rules</h2>"
$HTMLBody += Get-TransportRule | Select-Object Name, State | ConvertTo-Html -Fragment

$HTMLBody += "<h2>Send Connectors</h2>"
$HTMLBody += Get-SendConnector | Select-Object Name, 
    @{n='AddressSpaces';e={$_.AddressSpaces -join ', '}}, 
    Enabled, 
    @{n='SmartHosts';e={$_.SmartHosts -join ', '}} | ConvertTo-Html -Fragment

$HTMLBody += "<h2>Receive Connectors</h2>"
$HTMLBody += Get-ReceiveConnector | Select-Object Name, 
    Enabled, 
    @{n='PermissionGroups';e={$_.PermissionGroups -join ', '}} | ConvertTo-Html -Fragment

$HTMLBody += "<h2>Exchange Certificates</h2>"
$HTMLBody += Get-ExchangeCertificate | Select-Object @{n='Services';e={$_.Services -join ', '}}, Thumbprint, IsSelfSigned, NotAfter | ConvertTo-Html -Fragment

# ---------------------------------------------------------
# Virtual Directories
# ---------------------------------------------------------
Write-Host "Gathering Virtual Directories..." -ForegroundColor Cyan
$VDirData = @()
$VDirData += Get-OwaVirtualDirectory | Select-Object @{n='Service';e={'OWA'}}, Server, InternalUrl, ExternalUrl
$VDirData += Get-EcpVirtualDirectory | Select-Object @{n='Service';e={'ECP'}}, Server, InternalUrl, ExternalUrl
$VDirData += Get-WebServicesVirtualDirectory | Select-Object @{n='Service';e={'EWS'}}, Server, InternalUrl, ExternalUrl
$VDirData += Get-MapiVirtualDirectory | Select-Object @{n='Service';e={'MAPI'}}, Server, InternalUrl, ExternalUrl
$VDirData += Get-OabVirtualDirectory | Select-Object @{n='Service';e={'OAB'}}, Server, InternalUrl, ExternalUrl
$VDirData += Get-ActiveSyncVirtualDirectory | Select-Object @{n='Service';e={'EAS'}}, Server, InternalUrl, ExternalUrl
$HTMLBody += "<h2>Virtual Directories</h2>"
$HTMLBody += $VDirData | ConvertTo-Html -Fragment

# Assemble and output the Exchange HTML report
ConvertTo-Html -Head $CSS -Body $HTMLBody | Out-File $ReportPath
Write-Host "Exchange Report exported to: $ReportPath" -ForegroundColor Green
