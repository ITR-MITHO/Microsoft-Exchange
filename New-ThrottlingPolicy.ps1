New-ThrottlingPolicy -Name "Default Policy"
Set-ThrottlingPolicy "Default Policy" -EasMaxInactivityForDeviceCleanup 30

Foreach ($Mailbox in Get-Mailbox -RecipientTypeDetails UserMailbox -Resultsize Unlimited)
{

Set-Mailbox -identity $Mailbox.SamAccountName -ThrottlingPolicy "Default Policy"
    
}
