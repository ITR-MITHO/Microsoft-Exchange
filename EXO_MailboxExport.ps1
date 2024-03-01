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
$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
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
    $ArchiveInMB = "0"
}

$Statistics = Get-MailboxStatistics -Identity $Mailbox.SamAccountName | Select TotalItemSize, TotalDeletedItemSize
if ($Statistics) 
{
    $Size =  [math]::Round(([long]((($Statistics.TotalItemSize.Value -split "\(")[1] -split " ")[0] -split "," -join ""))/[math]::Pow(1024,3),3)
    $Deleted =  [math]::Round(([long]((($Statistics.TotalDeletedItemSize.Value -split "\(")[1] -split " ")[0] -split "," -join ""))/[math]::Pow(1024,3),3)
} 
else 
{
    $Size = "0"
    $Deleted = "0"
}
$Data = @{
    Username = $Mailbox.SamAccountName
    Name = $Mailbox.DisplayName
    Email = $Mailbox.PrimarySmtpAddress
    Type = $Mailbox.RecipientTypeDetails
    MailboxSize = $Size
    ArchiveSize = $ArchiveInMB
    Deleted = $Deleted
}   
$Results += New-Object PSObject -Property $Data
}
# Selecting the fields in a specific order instead of random.
$Results | Select Username, Name, Email, Type, MailboxSize, ArchiveSize, Deleted | 
Export-csv "$Home\Desktop\MailboxExport.csv" -NoTypeInformation -Encoding Unicode

Write-Host "
Find your .csv-file here: $Home\desktop\MailboxExport.csv" -ForegroundColor Green
