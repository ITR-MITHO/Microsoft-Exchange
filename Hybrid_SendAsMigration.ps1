<#
.SYNOPSIS
    Migrates on-premises SendAs permissions to Exchange Online.
.DESCRIPTION
    Finds explicit, non-self SendAs permissions on local mailboxes, maps trustees
    to their Primary SMTP addresses, and applies those permissions to the corresponding 
    mailboxes in Exchange Online.
.OUTPUTS
    $Home\Desktop\SendAs.csv - Full audit of on-premises SendAs permissions.
    $Home\Desktop\Failed.txt - Cloud synchronization failures.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

$CsvPath = Join-Path $home "Desktop\SendAs.csv"
$LogPath = Join-Path $home "Desktop\Failed.txt"

Write-Host "Gathering on-premises mailboxes and building lookup table..." -ForegroundColor Cyan
$Mailboxes = Get-Mailbox -ResultSize Unlimited

# Create a fast hashtable mapping SamAccountName/User domain suffix to PrimarySmtpAddress
$MailboxLookup = @{}
foreach ($M in $Mailboxes) {
    if (-not $MailboxLookup.ContainsKey($M.SamAccountName)) {
        $MailboxLookup[$M.SamAccountName] = $M.PrimarySmtpAddress
    }
}

$SendAsObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
Write-Host "Processing permissions..." -ForegroundColor Cyan
foreach ($Mailbox in $Mailboxes) {

    $SendAsPermissions = Get-ADPermission $Mailbox.DistinguishedName -ErrorAction SilentlyContinue | Where-Object {
        $_.ExtendedRights -like "*send*" -and 
        $_.IsInherited -eq $false -and 
        $_.User -notlike "*Self*" -and 
        $_.User -notlike "S-1-5-21*"
    }

    foreach ($Perm in $SendAsPermissions) {
        # Extract SamAccountName from 'DOMAIN\SamAccountName' or raw NT identity
        $UserDomain = $Perm.User -split '\\' | Select-Object -Last 1

        if ($MailboxLookup.ContainsKey($UserDomain)) {
            $UserEmail = $MailboxLookup[$UserDomain]
        } else {
            $UserEmail = $Perm.User # Fallback to original value if not an on-prem mailbox
        }

        $SendAsObjects.Add([PSCustomObject]@{
            MailboxSamAccountName = $Mailbox.SamAccountName
            MailboxDisplayName    = $Mailbox.DisplayName
            MailboxPrimarySMTP    = $Mailbox.PrimarySmtpAddress
            MailboxType           = $Mailbox.RecipientTypeDetails
            UserWithSendAs        = $UserEmail
        })
    }
}

$SendAsObjects | Export-Csv $CsvPath -NoTypeInformation -Encoding Unicode
Write-Host "On-prem audit saved to $CsvPath" -ForegroundColor Green

if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    Write-Warning "ExchangeOnlineManagement module missing. Attempting automated installation..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

try {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Exchange Online. Try to avoid running the script in Powershell ISE"
    break
}

Write-Host "Assigning SendAs permissions in Exchange Online..." -ForegroundColor Cyan
"User,Mailbox" | Out-File $LogPath -Encoding utf8

foreach ($Row in $SendAsObjects) {
    $Mailbox = $Row.MailboxPrimarySMTP
    $User    = $Row.UserWithSendAs

    try {
        Add-RecipientPermission -Identity $Mailbox -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "Successfully assigned SendAs: $User -> $Mailbox" -ForegroundColor Green
    } catch {
        Write-Host "Failed to assign SendAs: $User -> $Mailbox" -ForegroundColor Red
        "$User,$Mailbox" | Out-File $LogPath -Append -Encoding utf8
    }
}

Write-Host "Execution complete. Failures logged here: $LogPath" -ForegroundColor Yellow
