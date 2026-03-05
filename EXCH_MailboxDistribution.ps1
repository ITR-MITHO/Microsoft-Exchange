<#

Uses the attributes Username, Email and SizeInMB from the EXCH_MailboxExport.ps1 script. 
Distributes mailboxes to new databases based on the amount and desired size of each database, ensuring they are within the range of the desired database size.


#>

$DatabaseCount = 60
$DatabaseSizeMB = 256000    # 250GB
$ToleranceMB = 10000        # +10GB allowed

$TargetMax = $DatabaseSizeMB + $ToleranceMB
$Mailboxes = Import-Csv "$Home\desktop\Test1.csv"

# Sort mailboxes largest first (important for balancing)
$Mailboxes = $Mailboxes | Sort-Object SizeInMB -Descending
# Create database tracking
$Databases = @()

for ($i = 1; $i -le $DatabaseCount; $i++) {
    $databases += [PSCustomObject]@{
        Name = "DB$($i.ToString('000'))"
        SizeMB = 0
    }
}

$result = @()

foreach ($mb in $mailboxes) {

    $size = [int]$mb.SizeInMB

    # Try to find DB that stays under TargetMax
    $candidate = $databases |
        Where-Object { ($_.SizeMB + $size) -le $TargetMax } |
        Sort-Object SizeMB |
        Select-Object -First 1

    # If none fit under limit, choose DB with smallest overflow
    if (-not $candidate) {

        $candidate = $databases |
            Sort-Object { ($_.SizeMB + $size) - $TargetMax } |
            Select-Object -First 1

        Write-Warning "Mailbox $($mb.Email) ($size MB) exceeds target limit for all DBs. Assigned to $($candidate.Name)."
    }

    # Update DB size
    $candidate.SizeMB += $size

    # Store result
    $result += [PSCustomObject]@{
        Username = $mb.Username
        Email = $mb.Email
        DB = $candidate.Name
        SizeInMB = $size
    }
}

# Export mailbox → database mapping
$result | Export-Csv "$Home\Desktop\MailboxDatabaseMapping.csv" -NoTypeInformation

# Optional: show database size summary
$databases | Sort-Object Name | Format-Table Name,SizeMB
