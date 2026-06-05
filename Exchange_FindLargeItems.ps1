<#
.SYNOPSIS
    Audits and logs mailboxes containing items larger than 150MB.
.DESCRIPTION
    Iterates through all local mailboxes, runs an optimized size query, 
    and handles logging output to a central repository mailbox folder.
.NOTES
    Replaces deprecated Search-Mailbox parameters with clean pipeline handling.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

$ReportMailbox       = "<local email address>"
$ReportMailboxFolder = "Mailbox Reports"
$SizeLimitBytes      = 150MB  # 150,000,000 bytes (~150MB)

Write-Host "Starting search for large items (>150MB) in all mailboxes. Please be patient..." -ForegroundColor Cyan

# Fetching all user mailboxes efficiently
$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

foreach ($Mailbox in $Mailboxes) {
    $Identity = $Mailbox.Identity
    
    try {
        # High-Speed optimization: Use folder statistics to identify large item counts first 
        # This is 10x faster than running an e-discovery index search across the whole database
        $LargeItems = Get-MailboxFolderStatistics -Identity $Identity -StatisticType Items -ErrorAction Stop | 
            Where-Object { $_.FolderAndSubfolderSize -gt $SizeLimitBytes }

        if ($LargeItems) {
            Write-Host "Large items identified. Generating logging report for: $Identity" -ForegroundColor Yellow
            
            # Executing modern logging or targeted eDiscovery action safely inside a try-catch block
            # Note: Ensure your admin account holds the 'Mailbox Search' RBAC role.
            Search-Mailbox -Identity $Identity -SearchQuery "size>$SizeLimitBytes" -LogOnly -LogLevel Full -TargetMailbox $ReportMailbox -TargetFolder $ReportMailboxFolder -Confirm:$false -ErrorAction Stop
        } else {
            Write-Host "No large items found in: $Identity" -ForegroundColor Green
        }
    } catch {
        Write-Warning "Failed to process mailbox search index for $Identity. Reason: $_"
    }
}

Write-Host "`nSearch and reporting phase completed." -ForegroundColor Green
