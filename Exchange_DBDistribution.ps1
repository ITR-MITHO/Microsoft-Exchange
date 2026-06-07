<#
.SYNOPSIS
    Distributes mailboxes across a fixed number of databases to hit a target size baseline.
.DESCRIPTION
    Optimized to eliminate array copy overhead and heavy pipeline filtering inside loops.

    Use the Exchange_ExportMailboxInfo.ps1 script output as $csvpath to have Username, Email and size.
#>

# ---------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------
$DatabaseCount  = 60
$DatabaseSizeMB = 256000    # ~250 GB
$ToleranceMB    = 10240     # Exactly 10 GB
$TargetMax      = $DatabaseSizeMB + $ToleranceMB

$CsvPath        = "$Home\Desktop\Test1.csv"
$OutputPath     = "$Home\Desktop\MailboxDatabaseMapping.csv"

# ---------------------------------------------------------
# PREPARATION & FAST INITIALIZATION
# ---------------------------------------------------------
if (-not (Test-Path $CsvPath)) {
    Write-Error "Source CSV not found at $CsvPath"
    return
}

Write-Output "Importing and sorting mailboxes..."
$Mailboxes = Import-Csv $CsvPath | 
    Select-Object Username, Email, @{Name='SizeInMB'; Expression={[int]$_.SizeInMB}} | 
    Sort-Object SizeInMB -Descending
$Databases = New-Object System.Collections.Generic.List[PSCustomObject]
for ($i = 1; $i -le $DatabaseCount; $i++) {
    $Databases.Add([PSCustomObject]@{
        Name   = "DB$($i.ToString('000'))"
        SizeMB = 0
    })
}
$ResultList = New-Object System.Collections.Generic.List[PSCustomObject]

# ---------------------------------------------------------
# DISTRIBUTION LOOP
# ---------------------------------------------------------
Write-Output "Distributing $($Mailboxes.Count) mailboxes across $DatabaseCount databases..."
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($mb in $Mailboxes) {
    $size = $mb.SizeInMB
    $SelectedDB = $null

    # Finds the absolute emptiest database that can safely accommodate this mailbox under the Max limit
    $BestFitSize = [int]::MaxValue
    foreach ($db in $Databases) {
        $NewSize = $db.SizeMB + $size
        if ($NewSize -le $TargetMax) {
            if ($db.SizeMB -lt $BestFitSize) {
                $BestFitSize = $db.SizeMB
                $SelectedDB = $db
            }
        }
    }

    # Fallback if no database can hold it under TargetMax without overflowing
    if ($null -eq $SelectedDB) {
        $MinOverflow = [int]::MaxValue
        foreach ($db in $Databases) {
            $Overflow = ($db.SizeMB + $size) - $TargetMax
            if ($Overflow -lt $MinOverflow) {
                $MinOverflow = $Overflow
                $SelectedDB = $db
            }
        }
        Write-Warning "Mailbox $($mb.Email) ($size MB) forces all DBs over target limit. Assigned to $($SelectedDB.Name) (Overflow: $MinOverflow MB)."
    }

    # Update DB size inline
    $SelectedDB.SizeMB += $size
    $ResultList.Add([PSCustomObject]@{
        Username = $mb.Username
        Email    = $mb.Email
        DB       = $SelectedDB.Name
        SizeInMB = $size
    })
}

$Stopwatch.Stop()
Write-Output "Distribution completed in $($Stopwatch.Elapsed.TotalSeconds.ToString('F2')) seconds."

# ---------------------------------------------------------
# EXPORT & MULTI-ANGLE ANALYSIS
# ---------------------------------------------------------
$ResultList | Export-Csv $OutputPath -NoTypeInformation
Write-Output "Mapping exported to: $OutputPath"

# Metrics Breakdown (The Analytic View)
Write-Output "`n================ DATABASE BALANCING METRICS ================"
$FinalSizes = $Databases.SizeMB
$AverageSize = ($FinalSizes | Measure-Object -Average).Average
$MaxSize = ($FinalSizes | Measure-Object -Maximum).Maximum
$MinSize = ($FinalSizes | Measure-Object -Minimum).Minimum
$Spread = $MaxSize - $MinSize

Write-Output "Target Per DB : $DatabaseSizeMB MB"
Write-Output "Average DB Size: [$( [math]::Round($AverageSize / 1024, 2) ) GB] ($([math]::Round($AverageSize)) MB)"
Write-Output "Largest DB     : [$( [math]::Round($MaxSize / 1024, 2) ) GB] ($MaxSize MB)"
Write-Output "Smallest DB    : [$( [math]::Round($MinSize / 1024, 2) ) GB] ($MinSize MB)"
Write-Output "Max Spread     : [$( [math]::Round($Spread / 1024, 2) ) GB] ($Spread MB) variance between largest and smallest"
Write-Output "============================================================"

# Show summary table
$Databases | Sort-Object Name | Format-Table Name, @{Name="Size (GB)"; Expression={[math]::Round($_.SizeMB / 1024, 2)}}, SizeMB
