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
New-ADComputer -Name OwaAppProxy -AccountPassword (Read-Host 'Enter password' -AsSecureString) -Description 'Alternate Service Account credentials for Exchange' -Enabled:$True -SamAccountName OwaAppProxy
Set-ADComputer OwaAppProxy -add @{"msDS-SupportedEncryptionTypes"="28"}

# Setting up Kerberos authentication
cd $Exscripts
.\RollAlternateServiceAccountPassword.ps1 -ToSpecificServer localhost -GenerateNewPasswordFor $domain\OwaAppProxy$


# Verify that ASA-credentials are created
Get-ClientAccessServer localhost -IncludeAlternateServiceAccountCredentialStatus | Format-List Name, AlternateServiceAccountConfiguration


# Create SPN
setspn -S http/mail.$domain $domain\OwaAppProxy$

# Change OWA & ECP URL's to the same as written in the AppProxy

$URL = https://mail.$domain/ # REMEMBER TO CHANGE THIS URL BEFORE RUNNING.


Set-OwaVirtualDirectory -Identity "$env:COMPUTERNAME\owa (Default Web Site)" -InternalUrl "$URL/owa" -ExternalUrl "$URL/owa"
Set-EcpVirtualDirectory -Identity "$env:COMPUTERNAME\ecp (Default Web Site)" -InternalUrl "$URL/ecp" -ExternalUrl "$URL/ecp"
