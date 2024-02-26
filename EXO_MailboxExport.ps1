<#
The script will export the following information from all mailboxes:

                            SamAccountName
                            DisplayName
                            PrimarySmtpAddress
                            RecipientTypeDetails
                            TotalItemSize
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
$Statistics = Get-MailboxStatistics -Identity $Mailbox.SamAccountName
if ($Statistics) 
{
    $Size =  [math]::Round(([long]((($Statistics.totalitemsize.value -split "\(")[1] -split " ")[0] -split "," -join ""))/[math]::Pow(1024,3),3)
    $Deleted =  [math]::Round(([long]((($Statistics.TotalDeletedItemSize.value -split "\(")[1] -split " ")[0] -split "," -join ""))/[math]::Pow(1024,3),3)
} 
else 
{
    $Size = $null
    $Deleted = $null
}
$Data = @{
    Username = $Mailbox.SamAccountName
    Name = $Mailbox.DisplayName
    Email = $Mailbox.PrimarySmtpAddress
    Type = $Mailbox.RecipientTypeDetails
    Size = "$Size MB"
    Deleted = "$Deleted MB"
}   
$Results += New-Object PSObject -Property $Data
}
# Selecting the fields in a specific order instead of random.
$Results | Select Username, Name, Email, Type, Size, Total | 
Export-csv "C:\Users\mitho\OneDrive - IT Relation\Skrivebord\MailboxExport.csv" -NoTypeInformation -Encoding Unicode

Write-Host "
Find your .csv-file here: $Home\desktop\MailboxExport.csv" -ForegroundColor Green
