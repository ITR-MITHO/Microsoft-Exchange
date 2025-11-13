# Retention Tag Age Definitions
$RetentionPeriods = @{
	"30 Days"  = 30
    "90 Days"  = 90
	"180 Days" = 180
    "1 Year"   = 365
    "2 Years"  = 730
    "3 Years"  = 1095
	"4 Years"  = 1460
	"5 Years"  = 1825
}

foreach ($label in $RetentionPeriods.Keys) {
    $days = $RetentionPeriods[$label]

    # Delete Tags
    New-RetentionPolicyTag -Name "ITM8 - All - Delete - $label" `
        -Type All `
        -RetentionEnabled $true `
        -AgeLimitForRetention $days `
        -RetentionAction DeleteAndAllowRecovery `
        -Comment "Automatically deletes items older than $label (allowing recovery)."
		
	New-RetentionPolicyTag -Name "ITM8 - Inbox - Delete - $label" `
        -Type Inbox `
        -RetentionEnabled $true `
        -AgeLimitForRetention $days `
        -RetentionAction DeleteAndAllowRecovery `
        -Comment "Automatically deletes items older than $label (allowing recovery)."
		
	New-RetentionPolicyTag -Name "ITM8 - DeletedItems - Delete - $label" `
        -Type DeletedItems `
        -RetentionEnabled $true `
        -AgeLimitForRetention $days `
        -RetentionAction DeleteAndAllowRecovery `
        -Comment "Automatically deletes items older than $label (allowing recovery)."

    # Archive Tags
    New-RetentionPolicyTag -Name "ITM8 - All - Archive - $label" `
        -Type All `
        -RetentionEnabled $true `
        -AgeLimitForRetention $days `
        -RetentionAction MoveToArchive `
        -Comment "Automatically archives items older than $label."
		

}

# Retention Policies
foreach ($label in $RetentionPeriods.Keys) {
    # Delete policy
    New-RetentionPolicy -Name "ITM8 - All - Delete - $label" `
        -RetentionPolicyTagLinks "ITM8 - All - Delete - $label"

    # Archive policy
    New-RetentionPolicy -Name "ITM8 - All - Archive - $label" `
        -RetentionPolicyTagLinks "ITM8 - All - Archive - $label"
		
	    New-RetentionPolicy -Name "ITM8 - Inbox - Delete - $label" `
        -RetentionPolicyTagLinks "ITM8 - Inbox - Delete - $label"
		
	    New-RetentionPolicy -Name "ITM8 - DeletedItems - Delete - $label" `
        -RetentionPolicyTagLinks "ITM8 - DeletedItems - Delete - $label"
}

Write-Host "All retention tags and policies have been created successfully." -ForegroundColor Green
