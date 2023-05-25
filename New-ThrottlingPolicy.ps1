<#

Creates a new ThrottlingPolicy and assigns it to all mailboxes.
The policy deletes ActiveSync Devices that haven't synced in over 30 days.

#>

New-ThrottlingPolicy -Name "Default Throttling Policy"
Set-ThrottlingPolicy "Default Throttling Policy" -EasMaxInactivityForDeviceCleanup 30

Foreach ($Mailbox in Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails UserMailbox)
{

Set-Mailbox $Mailbox.SamAccountName -ThrottlingPolicy "Default Policy"
    
}