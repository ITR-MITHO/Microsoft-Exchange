<#

Creates RetentionPolicies and Tags that either Delete or Archive ALL items in a mailbox. 

#>
# Retention Tag Age Definitions
$RetentionPeriods = @{
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
    New-RetentionPolicyTag -Name "ITM8 - $label - Delete" `
        -Type All `
        -RetentionEnabled $true `
        -AgeLimitForRetention $days `
        -RetentionAction DeleteAndAllowRecovery `
        -Comment "Automatically deletes items older than $label (allowing recovery)."

    # Archive Tags
    New-RetentionPolicyTag -Name "ITM8 - $label - Archive" `
        -Type All `
        -RetentionEnabled $true `
        -AgeLimitForRetention $days `
        -RetentionAction MoveToArchive `
        -Comment "Automatically archives items older than $label."
}
# Retention Policies
foreach ($label in $RetentionPeriods.Keys) {
    # Delete policy
    New-RetentionPolicy -Name "ITM8 - $label - Delete" `
        -RetentionPolicyTagLinks "ITM8 - $label - Delete"
        
    # Archive policy
    New-RetentionPolicy -Name "ITM8 - $label - Archive" `
        -RetentionPolicyTagLinks "ITM8 - $label - Archive"
}
Write-Host "All retention tags and policies have been created successfully." -ForegroundColor Green
