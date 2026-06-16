<#
Run from on-prem Exchange. 
The script will prompt for O365 credentials to connect to Microsoft Graph to gather license information about all on-prem mailboxes.
#>

# Ensure Exchange Management Shell cmdlets are available
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

# Prerequisite Checks: Microsoft Graph Module Verification
$RequiredModules = @{
    'Microsoft.Graph.Authentication' = 'Connect-MgGraph'
    'Microsoft.Graph.Users'          = 'Get-MgUserLicenseDetail'
}

foreach ($Module in $RequiredModules.Keys) {
    $Cmdlet = $RequiredModules[$Module]
    if (-not (Get-Command $Cmdlet -ErrorAction SilentlyContinue)) {
        if (-not (Get-Module -ListAvailable $Module)) {
            Write-Host "Installing missing PowerShell Module: $Module. Please wait..." -ForegroundColor Yellow
            Install-Module $Module -Scope CurrentUser -AllowClobber -Force -Confirm:$false
        }
        Import-Module $Module -ErrorAction Stop
    }
}

# Connect to Graph API
try {
    Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Microsoft Graph. Check internet connectivity and administrative permissions."
    break
}

Write-Host "Gathering on-premises mailboxes..." -ForegroundColor Cyan
$Mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object SamAccountName, DisplayName, UserPrincipalName, PrimarySMTPAddress, RecipientTypeDetails

Write-Host "Bulk-fetching Active Directory user account properties..." -ForegroundColor Cyan
# Fetch all AD users into a fast lookup table to eliminate slow per-mailbox queries
$ADUsers = Get-ADUser -Filter * -Properties Enabled, LastLogonDate | Group-Object SamAccountName -AsHashTable -AsString

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Evaluating license status via Microsoft Graph..." -ForegroundColor Cyan
foreach ($Mailbox in $Mailboxes) {
    
    # Retrieve pre-cached AD attributes
    $ADUser = $ADUsers[$Mailbox.SamAccountName]
    $Enabled = if ($ADUser) { $ADUser.Enabled } else { $false }
    $LastLogonDate = if ($ADUser -and $ADUser.LastLogonDate) { $ADUser.LastLogonDate.ToString("dd-MM-yyyy") } else { "" }

    # Fetch User License from Graph using Primary SMTP (Matching your original logic)
    $RawLicenses = Get-MgUserLicenseDetail -UserId $Mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SkuPartNumber

    if (-not $RawLicenses) {
        $License = "No license"
    } else {
        # Join into a single string first, exactly like your working version
        $LicenseString = $RawLicenses -join ", "

        # Clean mapping using your exact string evaluation patterns
        $License = switch ($true) {
            ($LicenseString -like "*SPE_E3*")           { "Microsoft 365 E3" }
            ($LicenseString -like "*SPE_E5*")           { "Microsoft 365 E5" }
            ($LicenseString -like "*SPB*")              { "Microsoft Business Premium" }
            ($LicenseString -like "*EXCHANGESTANDARD*") { "Exchange Online Plan 1" }
            ($LicenseString -like "*EXCHANGEPREMIUM*")  { "Exchange Online Plan 2" }
            Default                                     { $LicenseString } # Fallback to raw string if unmapped
        }
    }

    $Results.Add([PSCustomObject]@{
        DisplayName = $Mailbox.DisplayName
        Username    = $Mailbox.SamAccountName
        Email       = $Mailbox.PrimarySMTPAddress
        Licens      = $License
        Type        = $Mailbox.RecipientTypeDetails
        Enabled     = $Enabled
        LastLogon   = $LastLogonDate
    })
}

$OutputPath = Join-Path $home "Desktop\Licenses.csv"
$Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding Unicode -Delimiter ";"

Write-Host "Export Completed, find your file here: $OutputPath" -ForegroundColor Green
