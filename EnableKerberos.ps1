<# 
Open Exchange Shell from an Exchange Server in a elevated mode
Microsoft Docs: https://docs.microsoft.com/en-us/exchange/architecture/client-access/kerberos-auth-for-load-balanced-client-access?view=exchserver-2019

#>
# Run this first to import the Active Directory module
Import-Module ActiveDirectory

# Verify that the SPN isn't already taken by another service.
Setspn -F -Q http/mail.domain.com

# Create a new ASA-credential named "EXCH-KERB" - Can be found in AD by searching for 'computers' when created
New-ADComputer -Name EXCH-KERB -AccountPassword (Read-Host 'Enter password' -AsSecureString) -Description 'Alternate Service Account credentials for Exchange' -Enabled:$True -SamAccountName EXCH-KERB
Set-ADComputer EXCH-KERB -add @{"msDS-SupportedEncryptionTypes"="28"}

# Setup Kerberos authentication by running the RollAlternateServiceAccountPassword script
cd $Exscripts
.\RollAlternateServiceAccountPassword.ps1 -ToSpecificServer localhost -GenerateNewPasswordFor domain.com\EXCH-KERB$

# Verify that the ASA-credential is created
Get-ClientAccessServer localhost -IncludeAlternateServiceAccountCredentialStatus | Format-List Name, AlternateServiceAccountConfiguration

# Create a SPN that matches your autodiscover URL
setspn -S http/mail.domain.com domain.com\EXCH-KERB$

# Verify that the SPN was created:
 setspn -L domain.com\EXCH-KERB$

# Enable kerberos authentication for Outlook clients:
Get-OutlookAnywhere -Server $env:localhost | Set-OutlookAnywhere -InternalClientAuthenticationMethod  Negotiate
Get-MapiVirtualDirectory -Server $env:localhost | Set-MapiVirtualDirectory -IISAuthenticationMethods OAuth,Kerberos
 
<#
Verify traffic is using Kerberos:
%ExchangeInstallPath%\Logging\HttpProxy\RpcHttp
Open the most recent log file, and then look for the word Negotiate
#>
