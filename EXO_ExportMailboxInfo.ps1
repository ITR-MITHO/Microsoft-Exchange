<#
.DESCRIPTION
    Compiles primary sizes, archive allocations, and deleted item (dumpster) metrics into a standardized inventory spreadsheet.
.OUTPUTS
    $Home\Desktop\MailboxExport.csv
#>

$CsvPath = Join-Path $home "Desktop\MailboxExport.csv"

Write-Host "Gathering target mailboxes..." -ForegroundColor Cyan
# Optimization: Filter out Discovery Mailboxes server-side to minimize network payload sizes
$Mailboxes = Get-Mailbox -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'"
$MailboxCount = $Mailboxes.Count
$Count = 1

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
Write-Host "Analyzing storage configurations. Please wait..." -ForegroundColor Yellow

foreach ($Mailbox in $Mailboxes) {
    $Sam = $Mailbox.SamAccountName
    
    # Progress Bar UI Control
    $Activity = 'Processing... [{0}/{1}]' -f $Count, $MailboxCount
    $Status   = 'Extracting mailbox statistics for: {0}' -f $Mailbox.DisplayName
    Write-Progress -Status $Status -Activity $Activity -PercentComplete (($Count / $MailboxCount) * 100)
    $PrimaryStats = Get-MailboxStatistics -Identity $Sam -ErrorAction SilentlyContinue
    
    $MailboxSizeMB = if ($PrimaryStats.TotalItemSize.Value) { $PrimaryStats.TotalItemSize.Value.ToMB() } else { 0 }
    $DeletedSizeMB = if ($PrimaryStats.TotalDeletedItemSize.Value) { $PrimaryStats.TotalDeletedItemSize.Value.ToMB() } else { 0 }

    $ArchiveSizeMB = "No Archive"
    if ($Mailbox.ArchiveStatus -ne "None") {
        $ArchiveStats = Get-MailboxStatistics -Identity $Sam -Archive -ErrorAction SilentlyContinue
        if ($ArchiveStats.TotalItemSize.Value) {
            $ArchiveSizeMB = $ArchiveStats.TotalItemSize.Value.ToMB()
        }
    }
    $Results.Add([PSCustomObject]@{
        Username             = $Mailbox.Alias
        Name                 = $Mailbox.DisplayName
        Email                = $Mailbox.PrimarySmtpAddress.ToString()
        Type                 = $Mailbox.RecipientTypeDetails
        MailboxSizeGB        = [math]::Round($MailboxSizeMB / 1024, 3) # Convert MB to GB cleanly via rounder
        ArchiveSizeGB        = if ($ArchiveSizeMB -is [num]) { [math]::Round($ArchiveSizeMB / 1024, 3) } else { $ArchiveSizeMB }
        TotalDeletedItemSize = [math]::Round($DeletedSizeMB / 1024, 3)
    })

    $Count++
}

Write-Progress -Activity "Processing..." -Completed
if ($Results.Count -gt 0) {
    $Results | Select-Object Username, Name, Email, Type, MailboxSizeGB, ArchiveSizeGB, TotalDeletedItemSize | 
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding Unicode
    
    Clear-Host
    Write-Host "Analysis complete! Report file compiled successfully to: $CsvPath" -ForegroundColor Green
} else {
    Write-Host "No mailbox data discovered to export." -ForegroundColor Yellow
}
