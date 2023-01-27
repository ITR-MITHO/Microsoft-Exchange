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
