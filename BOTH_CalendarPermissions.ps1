# Default UserMailbox Calendar Permissions
$User = 'Default'
$AccessRight = 'Reviewer'
Foreach ($Mailbox in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{
Try 
{ 
    Add-MailboxFolderPermission -Identity ($Mailbox.UserPrincipalName+":\$Calendar") -User $User -AccessRights $AccessRight -SendNotificationToUser $false -ErrorAction Stop
    Set-MailboxFolderPermission -Identity ($Mailbox.UserPrincipalName+":\$Calendar") -User $User -AccessRights $AccessRight -SendNotificationToUser $false -WarningAction Stop
    Write-Host "$User with $AccessRight added to $UserPrincipalName" -ForegroundColor Green
}
Catch
{
    Write-Warning "Permission already exist on Mailbox: $UserPrincipalName"
    Continue
}
	}
