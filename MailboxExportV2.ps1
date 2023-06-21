<#
V.2
The script will export the following information from all mailboxes:

                            SamAccountName
                            DisplayName
                            PrimarySmtpAddress
                            RecipientTypeDetails
                            DatabaseName
                            LastLogonTime
                            ADEnabled
                            TotalItemSize.To.MB
                            TotalDeletedItemSize.To.MB
                            Total

.NOTE
The attribute "Total" This field empty and used to add Size & Deleted together in this field inside Excel to determine the total size of a mailbox.

If you're having any issues with the script, please reach out to me.
https://github.com/ITR-MITHO

#>

# Checking permissions
$PMError = Test-Path $Home\desktop\PermissionIssue.txt
if ($PMError)
{
Remove-Item "$Home\desktop\PermissionIssue.txt" -Force
}
timeout 3
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
echo "Start PowerShell as an Administrator" > $Home\desktop\PermissionIssue.txt
Start $home\desktop\PermissionIssue.txt
Break
}

Add-PSSnapin *EXC*
Import-Module ActiveDirectory

$File = Test-path "$Home\desktop\MailboxExport.csv"
If ($File)
{

$Confirm = Read-Host "MailboxExport.csv already exists on your desktop. Do you want me to delete it for you? (Y/N)"
If ($Confirm -eq "Y")

{

Remove-Item "$Home\desktop\MailboxExport.csv" -Confirm:$false

}
    }

Clear-Host
$Mailboxes = Get-Mailbox -ResultSize Unlimited
$Results = @()

Write-Host "It is estimated to take 10-15 minutes for large organisations. Grab a nice cup of coffee :-)" -ForegroundColor Yellow
Sleep 5
Foreach ($Mailbox in $Mailboxes)
{

$Statistics = Get-MailboxStatistics -Identity $Mailbox.SamAccountName
$ADAtt = Get-ADUser -Identity $Mailbox.SamAccountName -Properties Enabled

if ($Statistics) 
{
    $Size = $Statistics.TotalItemSize.Value.ToMB()
    $Deleted = $Statistics.TotalDeletedItemSize.Value.ToMB()
} 
  
else 
{
    $Size = $null
    $Deleted = $null
}


If ($Statistics.LastLogonTime)
{

$LastLogon = $Statistics.LastlogonTime.ToString("dd-MM-yyyy")


}
Else
{

$LastLogon = ""

}

$results += [PSCustomObject]@{
    Username = $Mailbox.SamAccountName
    Name = $Mailbox.DisplayName
    Email = $Mailbox.PrimarySmtpAddress
    Type = $Mailbox.RecipientTypeDetails
    DB = $Statistics.DatabaseName
    LastLogon = $LastLogon
    ADEnabled = $ADAtt.Enabled
    Size = $Size
    Deleted = $Deleted
    Total = $null  # This field empty and used to add Size & Deleted together in this field inside Excel to determine the total size of a mailbox.

}
    }
        
# Selecting the fields in a specific order instead of random.
$Results | Select Username, Name, Email, Type, Size, Deleted, Total, DB, LastLogon, ADEnabled | 
Export-csv $home\Desktop\MailboxExport.csv -NoTypeInformation -Encoding Unicode

Write-Host "
            Find your .csv-file here: $Home\desktop\MailboxExport.csv
            
            
            For a export of full & send-as permissions use the following script: https://github.com/ITR-MITHO/Microsoft-Exchange/blob/main/FullAccessAndSendAs.ps1" -ForegroundColor Green
