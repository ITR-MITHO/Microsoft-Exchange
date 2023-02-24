<# 
IMPORTANT: This is for NO DAG setups. 
.HowTo

Change the "domain.com" to your appropriate domain-name. 
Make sure to run the commands in a elevated Exchange Shell.
Microsoft Docs: https://docs.microsoft.com/en-us/exchange/architecture/client-access/kerberos-auth-for-load-balanced-client-access?view=exchserver-2019
#>

# Verify that the SPN isn't already taken by another service.
setspn -F -Q http/mail.domain.com

# Creates a new ASA-credential named "OwaAppProxy" - Can be found in AD by searching for 'computers'
New-ADComputer -Name OwaAppProxy -AccountPassword (Read-Host 'Enter password' -AsSecureString) -Description 'Alternate Service Account credentials for Exchange' -Enabled:$True -SamAccountName OwaAppProxy
Set-ADComputer OwaAppProxy -add @{"msDS-SupportedEncryptionTypes"="28"}

# Setting up Kerberos authentication
cd $Exscripts
.\RollAlternateServiceAccountPassword.ps1 -ToSpecificServer EXCH01.domain.com -GenerateNewPasswordFor domain.com\OwaAppProxy$


# Verify that ASA-credentials are created
Get-ClientAccessServer EXCH01 -IncludeAlternateServiceAccountCredentialStatus | Format-List Name, AlternateServiceAccountConfiguration


# Create SPN
setspn -S http/mail.domain.com domain.com\OwaAppProxy$

# Change OWA & ECP URL's to the same as written in the AppProxy

$URL = https://mail.domain.com/ # REMEMBER TO CHANGE THIS URL BEFORE RUNNING.


Set-OwaVirtualDirectory -Identity "$env:COMPUTERNAME\owa (Default Web Site)" -InternalUrl "$URL/owa" -ExternalUrl "$URL/owa"
Set-EcpVirtualDirectory -Identity "$env:COMPUTERNAME\ecp (Default Web Site)" -InternalUrl "$URL/ecp" -ExternalUrl "$URL/ecp"
