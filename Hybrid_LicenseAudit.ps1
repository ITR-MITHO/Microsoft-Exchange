<#
.SYNOPSIS
    Correlates on-premises Exchange mailboxes with Microsoft Graph license states and AD account status.
.DESCRIPTION
    Gathers local mailboxes, pulls account properties from AD in bulk, retrieves assignment 
    details via Microsoft Graph, and exports the merged data to a CSV.
.OUTPUTS
    $Home\Desktop\Licenses.csv - Cleaned report with readable license names and account status.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

# 1. Prerequisite Checks: Microsoft Graph Module Verification
if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Host "Installing missing PowerShell Module: Microsoft.Graph. Please wait..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber -Force -Confirm:$false
    Write-Host "Microsoft Graph installed. Continuing execution..." -ForegroundColor Green
}

# Connect to Graph API
try {
    Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Microsoft Graph. Check internet connectivity and administrative permissions."
    break
}

# 2. Optimized High-Speed Data Gathering
Write-Host "Gathering on-premises mailboxes..." -ForegroundColor Cyan
$Mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object SamAccountName, DisplayName, UserPrincipalName, PrimarySMTPAddress, RecipientTypeDetails

Write-Host "Bulk-fetching Active Directory user account properties..." -ForegroundColor Cyan
# Fetching all AD users at once into a fast lookup table to eliminate per-mailbox loops
$ADUsers = Get-ADUser -Filter * -Properties Enabled, LastLogonDate | Group-Object SamAccountName -AsHashTable -AsString

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

# 3. Processing and Normalization Loop
Write-Host "Evaluating license status via Microsoft Graph..." -ForegroundColor Cyan
foreach ($Mailbox in $Mailboxes) {
    
    # Retrieve pre-cached AD attributes
    $ADUser = $ADUsers[$Mailbox.SamAccountName]
    $Enabled = if ($ADUser) { $ADUser.Enabled } else { $false }
    $LastLogonDate = if ($ADUser -and $ADUser.LastLogonDate) { $ADUser.LastLogonDate.ToString("dd-MM-yyyy") } else { "" }

    # Fetch User License from Graph using Primary SMTP
    $LicenseSkus = Get-MgUserLicenseDetail -UserId $Mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SkuPartNumber

    if (-not $LicenseSkus) {
        $FriendlyLicense = "No license"
    } else {
        # Translate SKUs using a clean switch statement instead of nested If/Else statements
        $TranslatedLicenses = foreach ($Sku in $LicenseSkus) {
            switch -wildcard ($Sku) {
                "*SPE_E3*"           { "Microsoft 365 E3" }
                "*SPE_E5*"           { "Microsoft 365 E5" }
                "*SPB*"              { "Microsoft Business Premium" }
                "*EXCHANGESTANDARD*" { "Exchange Online Plan 1" }
                "*EXCHANGEPREMIUM*"  { "Exchange Online Plan 2" }
                Default              { $Sku } # Fallback to the raw SKU code if unmapped
            }
        }
        $FriendlyLicense = $TranslatedLicenses -join ", "
    }

    $Results.Add([PSCustomObject]@{
        DisplayName = $Mailbox.DisplayName
        Username    = $Mailbox.SamAccountName
        Email       = $Mailbox.PrimarySMTPAddress.ToString()
        Licens      = $FriendlyLicense
        Type        = $Mailbox.RecipientTypeDetails
        Enabled     = $Enabled
        LastLogon   = $LastLogonDate
    })
}

# 4. Clean Export
$OutputPath = Join-Path $home "Desktop\Licenses.csv"
$Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding Unicode -Delimiter ";"

Write-Host "Export Completed, find your file here: $OutputPath" -ForegroundColor Green
