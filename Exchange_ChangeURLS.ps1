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

$DomainName = Read-Host "Enter URL prefix (e.g., mail.domain.com)"
$Autodiscover = Read-Host "Enter Autodiscover URL (e.g., autodiscover.domain.com)"
if ([string]::IsNullOrWhiteSpace($DomainName)) {
    Write-Error "A valid namespace domain name is required."
    break
}

if ([string]::IsNullOrWhiteSpace($Autodiscover)) {
    Write-Error "A valid namespace domain name is required."
    break
}

$URL = "https://$DomainName"
$TargetServer = (Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop).Name

Write-Host "Updating virtual directories on Server: $TargetServer using namespace: $URL" -ForegroundColor Cyan
$VDirs = @{
    "Set-OwaVirtualDirectory"         = @{Name="owa (Default Web Site)"; Path="owa"}
    "Set-MapiVirtualDirectory"        = @{Name="mapi (Default Web Site)"; Path="mapi"}
    "Set-OabVirtualDirectory"         = @{Name="oab (Default Web Site)"; Path="oab"}
    "Set-EcpVirtualDirectory"         = @{Name="ecp (Default Web Site)"; Path="ecp"}
    "Set-ActiveSyncVirtualDirectory"  = @{Name="Microsoft-Server-ActiveSync (Default Web Site)"; Path="Microsoft-Server-ActiveSync"}
}


foreach ($Cmdlet in $VDirs.Keys) {
    $VDir = $VDirs[$Cmdlet]
    $Identity = "$TargetServer\$($VDir.Name)"
    $UrlPath  = $VDir.Path
    
    try {
        Write-Host "Configuring $Cmdlet for identity: $Identity..." -ForegroundColor DarkCyan
        $Params = @{
            Identity    = $Identity
            InternalUrl = "$URL/$UrlPath"
            ExternalUrl = "$URL/$UrlPath"
        }
        & $Cmdlet @Params -ErrorAction Stop
    } catch {
        Write-Warning "Failed to configure $Identity. Reason: $_"
    }
}

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

try {
    Write-Host "Configuring Client Access Service Autodiscover URI to $Autodiscover..." -ForegroundColor DarkCyan
    
    Set-ClientAccessService -Identity $TargetServer -AutoDiscoverServiceInternalUri "https://$Autodiscover/Autodiscover/Autodiscover.xml" -ErrorAction Stop
} catch {
    Write-Error "Failed to update Client Access Service Autodiscover configuration: $_"
}
Write-Host "`nVirtual directory URL adjustment processing complete." -ForegroundColor Green
