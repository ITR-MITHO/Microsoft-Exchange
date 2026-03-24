$CSVPATH = "$Home\Desktop\MailboxExport.csv"
$Mailboxes = $Mailboxes = Get-Mailbox -ResultSize Unlimited | Where {$_.RecipientTypeDetails -NE "DiscoveryMailbox"}
$Results = @()
Write-Host "It is estimated to take 10-15 minutes for large organisations. Grab a nice cup of coffee :-)" -ForegroundColor Yellow
Foreach ($Mailbox in $Mailboxes)
{
$Archive = Get-MailboxStatistics -Identity $Mailbox.SamAccountName -Archive -ErrorAction SilentlyContinue | Select TotalItemSize
If ($Archive)
{
    $ArchiveInMB = [math]::Round(([long]((($Archive.TotalItemSize.Value -split "\(")[1] -split " ")[0] -split "," -join ""))/[math]::Pow(1024,3),3)
}
else
{
    $ArchiveInMB = "No Archive"
}

$Statistics = Get-MailboxStatistics -Identity $Mailbox.SamAccountName | Select TotalItemSize
if ($Statistics) 
{
    $SizeInMB =  [math]::Round(([long]((($Statistics.TotalItemSize.Value -split "\(")[1] -split " ")[0] -split "," -join ""))/[math]::Pow(1024,3),3)

} 
else 
{
    $SizeInMB = "0"

}

If ($Mailbox.IsDirSynced -Eq "True")
{
    $DirSync = "Yes"
}
Else
{
    $DirSync = "No"
}

$Data = @{
    Username = $Mailbox.Alias
    Name = $Mailbox.DisplayName
    Email = $Mailbox.PrimarySmtpAddress
    Type = $Mailbox.RecipientTypeDetails
    MailboxSize = $SizeInMB
    ArchiveSize = $ArchiveInMB
    Retention = $Mailbox.RetentionPolicy
    Forward = $Mailbox.ForwardingAddress
    DirSync = $DirSync
    MOA = ($Mailbox.EmailAddresses | Where-Object { $_ -match "^(SMTP|smtp):[^@]+@[A-Za-z0-9-]+\.onmicrosoft\.com$" } | ForEach-Object { ($_ -split ":")[1] }) -join ";"
    Proxy = $Mailbox.EmailAddresses
}   
$Results += New-Object PSObject -Property $Data
}
# Selecting the fields in a specific order instead of random.
$Results | Select-Object Username, Name, Email, Type, MailboxSize, ArchiveSize, Retention, Forward, DirSync, MOA, Proxy | 
Export-csv $CSVPATH -NoTypeInformation -Encoding Unicode
CLS
Write-Host "Find your .csv-file here: $CSVPATH" -ForegroundColor Green
