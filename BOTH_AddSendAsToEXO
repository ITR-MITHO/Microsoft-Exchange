<#
- RUN FROM ON-PREM EXCHANGE
You will be prompted for Exchange Online credentials and to install the Exchange Online PowerShell module if missing
The script will find all sendas permissions on-prem and then add those permissions to mailboxes in Exchange Online

If mailboxes shows in Failed.txt it most likely is due to an alias being mailbox@domain.local or an non-accepted domain.

.OUTPUT
Failed.txt - Tells you which mailbox a user wasn't added to with SendAs
SendAs.csv - A full list off all SendAs Permissions from on-prem

#>

Add-PSSnapin *EXC*
# Send-As Permissions
$SendAsObjects = @()
Write-Host "Gathering mailboxes..."
$Mailboxes = Get-Mailbox -ResultSize Unlimited
Write-Host "Gathering mailboxes...Done"

Foreach ($Mailbox in $Mailboxes) {
    Write-Host "Processing: " $Mailbox.DistinguishedName    
    $SendAs = Get-ADPermission $Mailbox.DistinguishedName | Where-Object {$_.ExtendedRights -like "*send*" -and $_.isinherited -like "*false*" -and $_.User -notlike "*Self*" -and $_.user -notlike "S-1-5-21*"} | Select-Object User, IdentityReference, AccessRights
    
    Foreach ($User in $SendAs) {
        $UserDomain = $User.User -split '\\' | Select-Object -Last 1
        $UserMailbox = Get-Mailbox $UserDomain -ErrorAction SilentlyContinue
        

        if ($UserMailbox) {
            $UserEmail = $UserMailbox.PrimarySmtpAddress
        } else {
            $UserEmail = $User.User
        }

        $SendAsObject = [PSCustomObject] @{
            MailboxSamAccountName = $Mailbox.SamAccountName
            MailboxDisplayName = $Mailbox.DisplayName
            MailboxPrimarySMTP = $Mailbox.PrimarySmtpAddress
            MailboxType = $Mailbox.RecipientTypeDetails
            UserWithSendAs = $UserEmail
        }

        $SendAsObjects += $SendAsObject
    }
}
$SendAsObjects | Select MailboxSamAccountName, MailboxDisplayName, MailboxPrimarySMTP, MailboxType, UserWithSendAs | Export-Csv $Home\Desktop\SendAs.csv -NoTypeInformation -Encoding Unicode

# Connect to Exchange Online
Try
{
Connect-ExchangeOnline -ErrorAction Stop
}
Catch
{
Write-Host "Failed to connect to Exchange Online. If your user requires Multi-factor authentication from this destination, it will not work." -ForegroundColor Red    
Write-Host "Try to run 'Connect-ExchangeOnline' manually, to see if it prompts for MFA
If it still fails, the module is missing. Use: Install-Module -Name ExchangeOnlineManagement" -ForeGroundColor Yellow
Break
}
Write-host "Connected to Exchange Online!" -ForeGroundColor Green
Write-Host "Assigning send-as permissions"

$Log = "$home\desktop\Failed.txt"
Write-Output "User,Mailbox" | Out-File $Log
$CSV = Import-csv $home\desktop\Sendas.csv
Foreach ($C in $CSV)
{

$Mailbox = $C.MailboxPrimarySMTP
$User = $C.UserWithSendAs
Try
{
Add-RecipientPermission -Identity $Mailbox -Trustee $User -AccessRights Sendas -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
}
Catch
{
    Write-Output "$User,$Mailbox" | Out-File $Log -Append
}
    }
