<#
.SYNOPSIS
Two files will be created on your desktop, named FullAccess.csv & Sendas.csv
It can take up to 5 minutes to complete the export, depending on the size of your organisation. 
#>

# Full-Access Permissions
Add-PSSnapin *EXC*
Import-Module ActiveDirectory
$FullObjects = @()
$Mailboxes = Get-Mailbox -ResultSize Unlimited

Foreach ($Mailbox in $Mailboxes) {
    $FullAccessUsers = Get-MailboxPermission $Mailbox.Identity | Where-Object {$_.isinherited -like "*false*" -and $_.User -notlike "*Self*" -and $_.user -notlike "S-1-5-21*"} | Select-Object User, IdentityReference, AccessRights

    Foreach ($User in $FullAccessUsers) {
        $FullObject = [PSCustomObject] @{
            MailboxSamAccountName = $Mailbox.SamAccountName
            MailboxDisplayName = $Mailbox.DisplayName
            MailboxPrimarySMTP = $Mailbox.PrimarySmtpAddress
            MailboxType = $Mailbox.RecipientTypeDetails
            UserWithFull = $User.User
            
        }
        $FullObjects += $FullObject
    }
}

$FullObjects | Select MailboxSamAccountName, MailboxDisplayName, MailboxPrimarySMTP, MailboxType, UserWithFull | Export-Csv $Home\Desktop\FullAccess.csv -NoTypeInformation -Encoding Unicode


# Send-As Permissions
$SendAsObjects = @()
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

Foreach ($Mailbox in $Mailboxes) {
    $SendAs = Get-ADPermission $Mailbox.Identity | Where-Object {$_.ExtendedRights -like "*send*" -and $_.isinherited -like "*false*" -and $_.User -notlike "*Self*" -and $_.user -notlike "S-1-5-21*"} | Select-Object User, IdentityReference, AccessRights

    Foreach ($User in $SendAs) {
        $SendAsObject = [PSCustomObject] @{
            MailboxSamAccountName = $Mailbox.SamAccountName
            MailboxDisplayName = $Mailbox.DisplayName
            MailboxPrimarySMTP = $Mailbox.PrimarySmtpAddress
            MailboxType = $Mailbox.RecipientTypeDetails
            UserWithSendAs = $User.User
            
        }
        $SendAsObjects += $SendAsObject
    }
}

$SendAsObjects | Select MailboxSamAccountName, MailboxDisplayName, MailboxPrimarySMTP, MailboxType, UserWithSendAs | Export-Csv $Home\Desktop\SendAs.csv -NoTypeInformation -Encoding Unicode
Write-Host "FullAccess.csv & SendAS.csv can be found here: $home\desktop" -ForegroundColor Green
