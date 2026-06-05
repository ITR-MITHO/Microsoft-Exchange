<#
.SYNOPSIS
    Exports volumetric storage statistics across organization mailboxes.
.DESCRIPTION
    Compiles primary sizes, archive allocations, and deleted item (dumpster) 
    metrics into a standardized inventory spreadsheet.
.OUTPUTS
    $Home\Desktop\MailboxExport.csv
#>

$CsvPath = Join-Path $home "Desktop\MailboxExport.csv"

# 1. Privileged Context Validation
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevated administrative privileges required. Exiting."
    break
}

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

Write-Host "Gathering target mailboxes..." -ForegroundColor Cyan
# Optimization: Filter out Discovery Mailboxes server-side to minimize network payload sizes
$Mailboxes = Get-Mailbox -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'"
$MailboxCount = $Mailboxes.Count
$Count = 1

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Analyzing storage configurations. Please wait..." -ForegroundColor Yellow

# 2. Storage Processing Engine Loop
foreach ($Mailbox in $Mailboxes) {
    $Sam = $Mailbox.SamAccountName
    
    # Progress Bar UI Control
    $Activity = 'Processing... [{0}/{1}]' -f $Count, $MailboxCount
    $Status   = 'Extracting mailbox statistics for: {0}' -f $Mailbox.DisplayName
    Write-Progress -Status $Status -Activity $Activity -PercentComplete (($Count / $MailboxCount) * 100)

    # Fetch primary stats (Consolidating calls to grab deleted item size natively)
    $PrimaryStats = Get-MailboxStatistics -Identity $Sam -ErrorAction SilentlyContinue
    
    # Safe object verification using native Byte Methods (Skips brittle string splitting)
    $MailboxSizeMB = if ($PrimaryStats.TotalItemSize.Value) { $PrimaryStats.TotalItemSize.Value.ToMB() } else { 0 }
    $DeletedSizeMB = if ($PrimaryStats.TotalDeletedItemSize.Value) { $PrimaryStats.TotalDeletedItemSize.Value.ToMB() } else { 0 }

    # Conditional evaluation for Archive states
    $ArchiveSizeMB = "No Archive"
    if ($Mailbox.ArchiveStatus -ne "None") {
        $ArchiveStats = Get-MailboxStatistics -Identity $Sam -Archive -ErrorAction SilentlyContinue
        if ($ArchiveStats.TotalItemSize.Value) {
            $ArchiveSizeMB = $ArchiveStats.TotalItemSize.Value.ToMB()
        }
    }

    # Fast literal casting directly into the type-safe generic collection list
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

# Clear visual progress engine status element
Write-Progress -Activity "Processing..." -Completed

# 3. Controlled File Output Phase
if ($Results.Count -gt 0) {
    $Results | Select-Object Username, Name, Email, Type, MailboxSizeGB, ArchiveSizeGB, TotalDeletedItemSize | 
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding Unicode
    
    Clear-Host
    Write-Host "Analysis complete! Report file compiled successfully to: $CsvPath" -ForegroundColor Green
} else {
    Write-Host "No mailbox data discovered to export." -ForegroundColor Yellow
}
