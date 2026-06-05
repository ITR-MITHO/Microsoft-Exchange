<#
.SYNOPSIS
    Configures internal and external URLs for Exchange Virtual Directories.
.DESCRIPTION
    Runs from an Exchange Management Shell context. Prompts for a namespace prefix
    and systematically updates OWA, MAPI, OAB, ECP, EWS, ActiveSync, Outlook Anywhere,
    and the Autodiscover Service Internal URI across the target server.
.OUTPUTS
    Console configuration log of successful or failed virtual directory bindings.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

# 1. Target Namespace & Context Extraction
$DomainName = Read-Host "Enter URL prefix (e.g., mail.domain.com)"
if ([string]::IsNullOrWhiteSpace($DomainName)) {
    Write-Error "A valid namespace domain name is required."
    break
}

$URL = "https://$DomainName"

# Dynamically target the local server using Exchange topology instead of env strings
$TargetServer = (Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop).Name

Write-Host "Updating virtual directories on Server: $TargetServer using namespace: $URL" -ForegroundColor Cyan

# Helper hash table to iterate through basic structural directory changes
$VDirs = @{
    "Set-OwaVirtualDirectory"         = "owa (Default Web Site)"
    "Set-MapiVirtualDirectory"        = "mapi (Default Web Site)"
    "Set-OabVirtualDirectory"         = "oab (Default Web Site)"
    "Set-EcpVirtualDirectory"         = "ecp (Default Web Site)"
    "Set-ActiveSyncVirtualDirectory"  = "Microsoft-Server-ActiveSync (Default Web Site)"
}

# 2. Process Standard Virtual Directories
foreach ($Cmdlet in $VDirs.Keys) {
    $VDirName = $VDirs[$Cmdlet]
    $Identity = "$TargetServer\$VDirName"
    
    try {
        Write-Host "Configuring $Cmdlet for identity: $Identity..." -ForegroundColor DarkCyan
        Invoke-Expression "$Cmdlet -Identity '$Identity' -InternalUrl '$URL/$($VDirName -split ' ')[0]' -ExternalUrl '$URL/$($VDirName -split ' ')[0]' -ErrorAction Stop"
    } catch {
        Write-Warning "Failed to configure virtual directory via $Cmdlet for $Identity. Reason: $_"
    }
}

# 3. Handle Special Path Rules (EWS & Outlook Anywhere)
try {
    Write-Host "Configuring Web Services (EWS)..." -ForegroundColor DarkCyan
    Set-WebServicesVirtualDirectory -Identity "$TargetServer\ews (Default Web Site)" -InternalUrl "$URL/EWS/Exchange.asmx" -ExternalUrl "$URL/EWS/Exchange.asmx" -ErrorAction Stop
} catch {
    Write-Warning "Failed to configure EWS: $_"
}

try {
    Write-Host "Configuring Outlook Anywhere..." -ForegroundColor DarkCyan
    Get-OutlookAnywhere -Server $TargetServer -ErrorAction SilentlyContinue | 
        Set-OutlookAnywhere -ExternalHostname $DomainName -InternalHostname $DomainName -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $true -DefaultAuthenticationMethod NTLM -ErrorAction Stop
} catch {
    Write-Warning "Failed to configure Outlook Anywhere: $_"
}

# 4. Modern Automated Autodiscover Remediation
try {
    $AutodiscoverUrl = "https://autodiscover.$($DomainName -replace '^mail\.')/autodiscover/autodiscover.xml"
    Write-Host "Configuring Client Access Service Autodiscover URI to $AutodiscoverUrl..." -ForegroundColor DarkCyan
    
    # Utilizing modern Exchange cmdlet instead of deprecated variations
    Set-ClientAccessService -Identity $TargetServer -AutoDiscoverServiceInternalUri $AutodiscoverUrl -ErrorAction Stop
    Write-Host "Autodiscover successfully updated." -ForegroundColor Green
} catch {
    Write-Error "Failed to update Client Access Service Autodiscover configuration: $_"
}

Write-Host "`nVirtual directory URL adjustment processing complete." -ForegroundColor Green
