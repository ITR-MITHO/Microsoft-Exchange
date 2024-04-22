# Finds out if a Mailbox have default MailboxPermissions or if "FullAccess" was added to it. 

$FullObjects = @()
$Mailboxes = Get-Mailbox -ResultSize Unlimited

Foreach ($Mailbox in $Mailboxes) {
    $FullAccessUsers = Get-MailboxPermission $Mailbox.Alias | Where-Object {$_.isinherited -like "*false*" -and $_.User -like "*@*"} | Select-Object User
    If ($FullAccessUsers)
    {
    $Custom = "Permissions added"
    }
    else
    {
    $Custom = "No permissions"
    }

$FullObject = [PSCustomObject] @{
            Alias = $Mailbox.Alias
            Display = $Mailbox.DisplayName
            Email = $Mailbox.PrimarySmtpAddress
            Type = $Mailbox.RecipientTypeDetails
            UserWithFull = $Custom
            
        }
        $FullObjects += $FullObject
    }
$FullObjects | Select Alias, Display, Email, Type, UserWithFull | Export-Csv "$home\desktop\CustomPermission.csv" -NoTypeInformation -Encoding Unicode
