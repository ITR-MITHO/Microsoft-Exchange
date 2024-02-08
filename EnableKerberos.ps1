<# 

Make sure to run the commands in a elevated Exchange Shell.
Microsoft Docs: https://docs.microsoft.com/en-us/exchange/architecture/client-access/kerberos-auth-for-load-balanced-client-access?view=exchserver-2019

#>

# Run this first
$domain = Read-host "Enter your domain name. (E.g. microsoft.com)"
Import-Module ActiveDirectory
Add-PSSnapin *EXC*

# Verify that the SPN isn't already taken by another service.

setspn -F -Q http/mail.$domain

# Creates a new ASA-credential named "OwaAppProxy" - Can be found in AD by searching for 'computers'
New-ADComputer -Name OwaAppProxy -AccountPassword (Read-Host 'Enter password' -AsSecureString) -Description 'Alternate Service Account credentials for Exchange' -Enabled:$True -SamAccountName EXCHKERB
Set-ADComputer EXCHKERB -add @{"msDS-SupportedEncryptionTypes"="28"}

# Setting up Kerberos authentication
cd $Exscripts
.\RollAlternateServiceAccountPassword.ps1 -ToSpecificServer localhost -GenerateNewPasswordFor $domain\EXCHKERB$


# Verify that ASA-credentials are created
Get-ClientAccessServer localhost -IncludeAlternateServiceAccountCredentialStatus | Format-List Name, AlternateServiceAccountConfiguration


# Create SPN
setspn -S http/mail.$domain $domain\EXCHKERB$

# Verify SPN:
 
setspn -L $Domain\EXCHKERB$

# Enable kerberos for Outlook clients:
Get-OutlookAnywhere -Server $env:localhost | Set-OutlookAnywhere -InternalClientAuthenticationMethod  Negotiate
Get-MapiVirtualDirectory -Server $env:localhost | Set-MapiVirtualDirectory -IISAuthenticationMethods Ntlm,Negotiate
 
<#
Verify traffic is using Kerberos:
%ExchangeInstallPath%\Logging\HttpProxy\RpcHttp
Open the most recent log file, and then look for the word Negotiate
#>
