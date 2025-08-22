<#
V.2
The script will export the following information from all mailboxes:

                            SamAccountName
                            DisplayName
                            PrimarySmtpAddress
                            UserPrincipalName
                            RecipientTypeDetails
                            DatabaseName
                            LastLogonTime
                            RetentionPolicy
                            ADEnabled
                            TotalItemSize.To.MB
                            ArchiveSize.To.MB
#>

$CSVPATH = "$Home\Desktop\MailboxExport.csv"

# Checking permissions
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
Write-Host "Start PowerShell as an Administrator" -ForeGroundColor Yellow
Break
}

# Adding PowerShell Modules
Add-PSSnapin *EXC*
Import-Module ActiveDirectory

# Gathering mailbox information
$Mailboxes = Get-Mailbox -ResultSize Unlimited
$MailboxCount = ($Mailboxes | Measure-Object).count
$Count = 1
$Results = @()
Foreach ($Mailbox in $Mailboxes)
{
    # Status bar while running
    $DisplayName = ('{0} ({1})' -f $Mailbox.DisplayName, $Mailbox.Name)
    $Activity = ('Working... [{0}/{1}]' -f $count, $MailboxCount)
    $Status = ('Getting mailbox information: {0}' -f $DisplayName)
    Write-Progress -Status $Status -Activity $Activity -PercentComplete (($Count / $MailboxCount) * 100)

$Statistics = Get-MailboxStatistics -Identity $Mailbox.SamAccountName | Select-Object TotalItemSize, TotalDeletedItemSize, LastLogonTime, DataBase
$ADAtt = Get-ADUser -Identity $Mailbox.SamAccountName -Properties Enabled, UserPrincipalName
$ArchiveSize = Get-MailboxStatistics -Identity $Mailbox.SamAccountName -Archive -ErrorAction SilentlyContinue | Select-Object TotalItemSize

If ($ArchiveSize)
{
    $ArchiveInMB = $ArchiveSize.TotalItemSize.Value.ToMB()
}
Else
{
    $ArchiveInMB = "No Archive"
}

If ($Statistics) 
{
    $Size = $Statistics.TotalItemSize.Value.ToMB()
} 
Else 
{
    $Size = "0"
}
If ($Statistics.LastLogonTime)
{
$LastLogon = $Statistics.LastlogonTime.ToString("dd-MM-yyyy")
}
Else
{
$LastLogon = $null
}

$Results += [PSCustomObject]@{
    Username = $Mailbox.SamAccountName
    Name = $Mailbox.DisplayName
    Email = $Mailbox.PrimarySmtpAddress
    UPN = $ADAtt.UserPrincipalName
    Type = $Mailbox.RecipientTypeDetails
    DB = $Statistics.Database
    LastLogon = $LastLogon
    ADEnabled = $ADAtt.Enabled
    SizeInMB = $Size
    ArchiveInMB = $ArchiveInMB
    Retention = $Mailbox.RetentionPolicy

}
$Count++ # End of status bar
    }

# Select-Object in a specific order instead of random.
$Results | Select-Object Username, Name, Email, UPN, Type, SizeInMB, ArchiveInMB, Retention, DB, LastLogon, ADEnabled | 
Export-csv $CSVPATH -NoTypeInformation -Encoding Unicode
Write-Host "Find your exported data here: $CSVPATH" -ForeGroundColor Green
