<#
.SYNOPSIS
    Exports both FullAccess and Send-As explicit permissions across all mailboxes.
.DESCRIPTION
    Runs a single optimization sweep over on-premises mailboxes to collect explicit, 
    non-inherited delegation mappings for audit verification.
.OUTPUTS
    $Home\Desktop\FullAccess.csv
    $Home\Desktop\SendAs.csv
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

Write-Host "Gathering target on-premises mailboxes..." -ForegroundColor Cyan
$Mailboxes = Get-Mailbox -ResultSize Unlimited

$FullAccessObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
$SendAsObjects     = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Auditing access and delegation states..." -ForegroundColor Cyan
foreach ($Mailbox in $Mailboxes) {
    
    $FullAccessPermissions = Get-MailboxPermission -Identity $Mailbox.Identity -ErrorAction SilentlyContinue | Where-Object {
        $_.IsInherited -eq $false -and 
        $_.User -notlike "*Self*" -and 
        $_.User -notlike "S-1-5-21*"
    }

    foreach ($Perm in $FullAccessPermissions) {
        $FullAccessObjects.Add([PSCustomObject]@{
            MailboxSamAccountName = $Mailbox.SamAccountName
            MailboxDisplayName    = $Mailbox.DisplayName
            MailboxPrimarySMTP    = $Mailbox.PrimarySmtpAddress
            MailboxType           = $Mailbox.RecipientTypeDetails
            UserWithFull          = $Perm.User
        })
    }

    $SendAsPermissions = Get-ADPermission -Identity $Mailbox.Identity -ErrorAction SilentlyContinue | Where-Object {
        $_.ExtendedRights -like "*send*" -and 
        $_.IsInherited -eq $false -and 
        $_.User -notlike "*Self*" -and 
        $_.User -notlike "S-1-5-21*"
    }

    foreach ($Perm in $SendAsPermissions) {
        $SendAsObjects.Add([PSCustomObject]@{
            MailboxSamAccountName = $Mailbox.SamAccountName
            MailboxDisplayName    = $Mailbox.DisplayName
            MailboxPrimarySMTP    = $Mailbox.PrimarySmtpAddress
            MailboxType           = $Mailbox.RecipientTypeDetails
            UserWithSendAs        = $Perm.User
        })
    }
}

$FullPath   = Join-Path $home "Desktop\FullAccess.csv"
$SendAsPath = Join-Path $home "Desktop\SendAs.csv"

$FullAccessObjects | Export-Csv -Path $FullPath -NoTypeInformation -Encoding Unicode
$SendAsObjects     | Export-Csv -Path $SendAsPath -NoTypeInformation -Encoding Unicode

Write-Host "`nExport complete! Files generated successfully:" -ForegroundColor Green
Write-Host " -> $FullPath" -ForegroundColor DarkGreen
Write-Host " -> $SendAsPath" -ForegroundColor DarkGreen
