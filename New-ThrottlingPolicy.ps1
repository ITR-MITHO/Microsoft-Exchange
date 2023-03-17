<#

Creates a new ThrottlingPolicy and assigns it to all mailboxes.
The policy deletes ActiveSync Devices that haven't synced in over 30 days.

#>

New-ThrottlingPolicy -Name "Default Policy"
Set-ThrottlingPolicy "Default Policy" -EasMaxInactivityForDeviceCleanup 30

Foreach ($Mailbox in Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails UserMailbox)
{

Set-Mailbox -identity $Mailbox.SamAccountName -ThrottlingPolicy "Default Policy"
    
}
