<#
.SYNOPSIS
    Validates Cloud mailboxes against On-Premises Active Directory and provisions missing Remote Mailboxes.
.DESCRIPTION
    Runs from an elevated on-premises Exchange management shell. Connects to Exchange Online, 
    identifies objects missing their local hybrid identity, and provisions them with the correct TargetAddress.
.OUTPUTS
    $Home\Desktop\RemoteMissing.csv - Audit log of mailboxes missing locally.
    $Home\Desktop\RemoteLog.csv     - Status report of provisioning actions.
#>

# 1. Enforcement & Prerequisites Check
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Elevated shell required. Please run this script as an Administrator."
    break
}

if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    Write-Error "The ExchangeOnlineManagement module is missing. Run 'Install-Module ExchangeOnlineManagement' first."
    break
}

# Ensure local Exchange context is active
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

# 2. Connect to Exchange Online and Collect Target Profiles
try {
    Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
} catch {
    Write-Error "Failed to authenticate to Exchange Online. Verify internet and account rights."
    break
}

$Domain = Read-Host "Enter Target Routing Tenant Domain (e.g., contoso.mail.onmicrosoft.com)"
if ([string]::IsNullOrWhiteSpace($Domain)) {
    Write-Error "A valid target routing domain is required."
    Disconnect-ExchangeOnline -Confirm:$false
    break
}

Write-Host "Fetching user mailboxes from Exchange Online..." -ForegroundColor Cyan
$EXOMailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | 
    Select-Object Alias, PrimarySmtpAddress, EmailAddresses

# Graceful Cloud Session Termination
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Cloud extraction finished. Processing local Directory..." -ForegroundColor Green

# 3. Cache On-Premises Directory Info for Speed
Write-Host "Indexing on-premises recipients to memory..." -ForegroundColor Cyan
# Build a fast index hash table of all local recipients to bypass inner-loop Get-Recipient executions
$LocalRecipients = @{}
Get-Recipient -ResultSize Unlimited | ForEach-Object {
    if ($_.Alias) { $LocalRecipients[$_.Alias.ToLower()] = $_.PrimarySmtpAddress }
}

$MissingObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
$TransactionLog = [System.Collections.Generic.List[PSCustomObject]]::new()

# 4. Perform Fast Memory Evaluation & Remediation
foreach ($Exo in $EXOMailboxes) {
    $Alias = $Exo.Alias
    $PrimarySmtp = $Exo.PrimarySmtpAddress.ToString()
    $RemoteRoutingAddress = "$Alias@$Domain"
    
    # Flatten the proxy addresses cleanly into a semi-colon separated string
    $ProxyAddressesString = ($Exo.EmailAddresses | ForEach-Object { $_.ProxyAddressString }) -join ";"

    # Check if the alias exists in our memory map instead of calling Active Directory repeatedly
    if (-not $LocalRecipients.ContainsKey($Alias.ToLower())) {
        
        # Track that it is missing locally
        $MissingObjects.Add([PSCustomObject]@{
            Alias              = $Alias
            PrimarySMTPAddress = $PrimarySmtp
            RemoteRouting      = $RemoteRoutingAddress
            EmailAddresses     = $ProxyAddressesString
        })

        # Remediate locally on-premises
        try {
            Write-Host "Provisioning Remote Mailbox for: $PrimarySmtp" -ForegroundColor Handled
            
            # Executing enablement command locally
            Enable-RemoteMailbox -Identity $PrimarySmtp -RemoteRoutingAddress $RemoteRoutingAddress -ErrorAction Stop
            
            $TransactionLog.Add([PSCustomObject]@{
                Email  = $PrimarySmtp
                Status = "SuccessfullyUpdated"
            })
        } catch {
            Write-Host "Failed to provision Remote Mailbox for: $PrimarySmtp. Reason: $_" -ForegroundColor Red
            $TransactionLog.Add([PSCustomObject]@{
                Email  = $PrimarySmtp
                Status = "FailedToUpdate"
            })
        }
    }
}

# 5. Output Reporting Documents
$MissingPath = Join-Path $home "Desktop\RemoteMissing.csv"
$LogPath     = Join-Path $home "Desktop\RemoteLog.csv"

if ($MissingObjects.Count -gt 0) {
    $MissingObjects | Export-Csv -Path $MissingPath -NoTypeInformation -Encoding Unicode -Delimiter ";"
    Write-Host "List of items missing locally exported to: $MissingPath" -ForegroundColor Green
} else {
    Write-Host "Perfect sync! No objects were missing on-premises." -ForegroundColor Green
}

if ($TransactionLog.Count -gt 0) {
    $TransactionLog | Export-Csv -Path $LogPath -NoTypeInformation -Encoding Unicode
    Write-Host "Action transaction log exported to: $LogPath" -ForegroundColor Yellow
}
