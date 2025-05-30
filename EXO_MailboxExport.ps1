<#
The script will export the following information from all mailboxes:

                            SamAccountName
                            DisplayName
                            PrimarySmtpAddress
                            RecipientTypeDetails
                            TotalItemSize
                            ArchiveSize
                            TotalDeletedItemSize


.NOTE

If you're having any issues with the script, please reach out to me.
https://github.com/ITR-MITHO

#>

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
$Data = @{
    Username = $Mailbox.Alias
    Name = $Mailbox.DisplayName
    Email = $Mailbox.PrimarySmtpAddress
    Type = $Mailbox.RecipientTypeDetails
    MailboxSize = $SizeInMB
    ArchiveSize = $ArchiveInMB
}   
$Results += New-Object PSObject -Property $Data
}
# Selecting the fields in a specific order instead of random.
$Results | Select-Object Username, Name, Email, Type, MailboxSize, ArchiveSize | 
Export-csv $CSVPATH -NoTypeInformation -Encoding Unicode
CLS
Write-Host "Find your .csv-file here: $CSVPATH" -ForegroundColor Green
