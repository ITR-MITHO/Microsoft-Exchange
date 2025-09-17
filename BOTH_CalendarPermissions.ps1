# Default UserMailbox Calendar Permissions
$User = 'Default'
$AccessRight = 'Reviewer'
Foreach ($Mailbox in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
{
$UserPrincipalName = $Mailbox.UserPrincipalName
$Calendar = (Get-MailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -FolderScope Calendar | Where { $_.FolderType -eq 'Calendar'}).Name
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
