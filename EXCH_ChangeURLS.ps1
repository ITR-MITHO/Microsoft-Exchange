Add-PSSnapin *EXC*
$DomainName = Read-Host "Enter url prefix like this: mail.domain.com"
$URL = "https://$DomainName"
                                                                                           

Set-OwaVirtualDirectory -Identity "$env:COMPUTERNAME\owa (Default Web Site)" -InternalUrl "$URL/owa" -ExternalUrl "$URL/owa"
Set-MapiVirtualDirectory -Identity "$env:COMPUTERNAME\mapi (Default Web Site)" -InternalUrl "$URL/mapi" -ExternalUrl "$URL/mapi"
Set-OabVirtualDirectory -Identity "$env:COMPUTERNAME\oab (Default Web Site)" -InternalUrl "$URL/oab" -ExternalUrl "$URL/oab"
Set-EcpVirtualDirectory -Identity "$env:COMPUTERNAME\ecp (Default Web Site)" -InternalUrl "$URL/ecp" -ExternalUrl "$URL/ecp"
Set-WebServicesVirtualDirectory -Identity "$env:COMPUTERNAME\ews (Default Web Site)" -InternalUrl "$URL/EWS/Exchange.asmx" -ExternalUrl "$URL/EWS/Exchange.asmx"
Set-ActiveSyncVirtualDirectory -Identity "$env:COMPUTERNAME\Microsoft-Server-ActiveSync (Default web site)" -InternalUrl "$URL/Microsoft-Server-ActiveSync" -ExternalUrl "$URL/Microsoft-Server-ActiveSync"
Get-OutlookAnywhere -Server "$env:COMPUTERNAME" | Set-OutlookAnywhere -ExternalHostname "$DomainName" -InternalHostname "$DomainName" -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $true -DefaultAuthenticationMethod NTLM


Write-Host "Urls changed for OWA, MAPI, OAB, ECP, EWS, EAS and Outlook anywhere" -ForegroundColor Green
Write-Host "
Remember to manually change the autodiscover URL
Set-ClientAccessServer -Identity $env:computername -AutoDiscoverServiceInternalUri MAIL.DOMAIN.COM/autodiscover/autodiscover.xml" -ForegroundColor Yellow
