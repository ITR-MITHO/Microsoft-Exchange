$DatabaseCount = 60
$DatabaseSizeMB = 256000     # 250GB
$ToleranceMB = 15000         # +/- 15GB

$TargetMax = $DatabaseSizeMB + $ToleranceMB

$mailboxes = Import-Csv "$Home\Desktop\Test1.csv"

# Sort largest first (important for good balancing)
$mailboxes = $mailboxes | Sort-Object SizeInMB -Descending

# Create database tracking
$databases = @()

for ($i = 1; $i -le $DatabaseCount; $i++) {
    $databases += [PSCustomObject]@{
        Name = "DB$i"
        SizeMB = 0
    }
}

$result = @()

foreach ($mb in $mailboxes) {

    $size = [int]$mb.SizeInMB

    # Find DB that stays under target max
    $candidate = $databases |
        Where-Object { ($_.SizeMB + $size) -le $TargetMax } |
        Sort-Object SizeMB |
        Select-Object -First 1

    # If none available choose smallest DB
    if (-not $candidate) {
        $candidate = $databases | Sort-Object SizeMB | Select-Object -First 1
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

# Export mapping file
$result | Export-Csv "$Home\Desktop\MailboxDatabaseMapping.csv" -NoTypeInformation
