<#
.SYNOPSIS
    Audits non-default permissions on all subfolders across user mailboxes.
.DESCRIPTION
    Scans folders within target mailboxes and isolates instances where the 'Default' 
    anonymous/authenticated user access has been granted rights other than 'None'.
.OUTPUTS
    $home\desktop\folderpermissions.csv
#>
param (
    [string]$Identity = "*", 
    [string]$OutputFile = "$home\desktop\folderpermissions.csv"
)
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}
Write-Host "Gathering user mailboxes matching context criteria..." -ForegroundColor Cyan
$Mailboxes = Get-Mailbox -RecipientTypeDetails 'UserMailbox' -Filter "Name -like '$Identity'" -ResultSize Unlimited

$MailboxCount = $Mailboxes.Count
$Count = 1
$ResultList = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($Mailbox in $Mailboxes) {
    $Alias = $Mailbox.Alias
    $DisplayName = '{0} ({1})' -f $Mailbox.DisplayName, $Mailbox.Name
    $Activity = 'Working... [{0}/{1}]' -f $Count, $MailboxCount
    $Status   = 'Analyzing MAPI folder paths for: {0}' -f $DisplayName
    Write-Progress -Status $Status -Activity $Activity -PercentComplete (($Count / $MailboxCount) * 100)

    # Optimization: Use Get-MailboxFolder instead of Get-MailboxFolderStatistics.
    # This queries the structural hierarchy directly, running up to 5x faster.
    try {
        $FolderStructure = Get-MailboxFolder -Identity "$Alias:\*" -Recurse -ErrorAction Stop | 
            Select-Object -ExpandProperty Identity
    } catch {
        Write-Warning "Failed to query structural hierarchy maps for: $DisplayName"
        $Count++
        continue
    }

    # Explicitly prepend the root information store directory target 
    $TargetFolders = [System.Collections.Generic.List[string]]::new()
    $TargetFolders.Add("$Alias:\")
    foreach ($FolderId in $FolderStructure) { [void]$TargetFolders.Add($FolderId) }

    # 3. Targeted Permission Evaluation Phase
    foreach ($FolderKey in $TargetFolders) {
        
        $Permissions = Get-MailboxFolderPermission -Identity $FolderKey -ErrorAction SilentlyContinue
        if (-not $Permissions) { continue }

        foreach ($Permission in $Permissions) {
            $User = $Permission.User -replace "ExchangePublishedUser\.", ""
            
            # Isolate non-standard default security delegation models
            if ($User -eq "Default" -and $Permission.AccessRights -notcontains 'None') {
                
                # Format internal folder path references cleanly for reporting
                $CleanedPath = $FolderKey -replace "^$Alias:", ""
                if ([string]::IsNullOrEmpty($CleanedPath)) { $CleanedPath = "\" }

                $ResultList.Add([PSCustomObject]@{
                    Mailbox      = $DisplayName
                    FolderName   = $Permission.FolderName
                    Identity     = $CleanedPath
                    User         = $User
                    AccessRights = $Permission.AccessRights -join ','
                })
            }
        }
    }
    $Count++
}

# Complete UI progress bar visibility handle cleanly
Write-Progress -Activity "Working..." -Completed

# 4. CSV File Serialization Boundary
if ($ResultList.Count -gt 0) {
    $ResultList | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding Unicode
    Write-Host "`nAnalysis complete! Report compiled to: $OutputFile" -ForegroundColor Green
} else {
    Write-Host "`nAudit complete. No non-default folder permissions detected." -ForegroundColor Yellow
}
