New-ThrottlingPolicy -Name "Default Policy"
Set-ThrottlingPolicy "Default Policy" -EasMaxInactivityForDeviceCleanup 30

Foreach ($Mailbox in Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails UserMailbox)
{

Set-Mailbox -identity $Mailbox.SamAccountName -ThrottlingPolicy "Default Policy"
    
}
