<#
.SYNOPSIS
    Generates an inventory report of all ActiveSync device partnerships.
.DESCRIPTION
    Gathers mobile device specifications and synchronization statistics from 
    all mailboxes with active synchronization partnerships.
.OUTPUTS
    $home\desktop\EASDevices.csv
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

$CurrentDate = Get-Date
$ExportPath  = Join-Path $home "Desktop\EASDevices.csv"
# Define an explicit cutoff if you want to filter out old devices (e.g., 0 to show all devices)
$AgeFilterDays = 0 

Write-Host "Gathering ActiveSync mailbox profiles..." -ForegroundColor Cyan

# Optimization: Bulk capture mailboxes with active partnerships server-side to skip slow clientside Where-Object parsing
$EasMailboxes = Get-CASMailbox -ResultSize Unlimited -Filter "HasActiveSyncDevicePartnership -eq `$true" | 
    Select-Object Identity, DisplayName

if ($EasMailboxes.Count -eq 0) {
    Write-Host "No active ActiveSync device partnerships discovered in the environment." -ForegroundColor Yellow
    break
}

Write-Host "Building primary SMTP directory address index..." -ForegroundColor Cyan
# Pre-cache PrimarySmtpAddress to skip running Get-Mailbox user-by-user inside the loop
$MailboxLookup = @{}
Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $MailboxLookup[$_.Identity.ToString()] = $_.PrimarySmtpAddress.ToString()
}

$ReportCollection = [System.Collections.Generic.List[PSCustomObject]]::new()
$TotalCount = $EasMailboxes.Count
$CurrentIndex = 1

Write-Host "Analyzing active device partnerships. Please wait..." -ForegroundColor Yellow
foreach ($Mailbox in $EasMailboxes) {
    $MailboxId = $Mailbox.Identity.ToString()
    
    $Activity = 'Processing... [{0}/{1}]' -f $CurrentIndex, $TotalCount
    $Status   = 'Extracting active mobile partnerships for: {0}' -f $Mailbox.DisplayName
    Write-Progress -Status $Status -Activity $Activity -PercentComplete (($CurrentIndex / $TotalCount) * 100)

    $DeviceStats = @(Get-ActiveSyncDeviceStatistics -Mailbox $MailboxId -ErrorAction SilentlyContinue)
    $PrimarySmtp = if ($MailboxLookup.ContainsKey($MailboxId)) { $MailboxLookup[$MailboxId] } else { "Unknown" }

    foreach ($Device in $DeviceStats) {
        $LastAttempt = $Device.LastSyncAttemptTime
        
        if ($null -eq $LastAttempt) {
            $SyncDays = "Never"
        } else {
            $SyncDays = ($CurrentDate - $LastAttempt).Days
        }

        # Apply standardized age filtering evaluation rules safely
        if ($SyncDays -eq "Never" -or $SyncDays -ge $AgeFilterDays) {
            
            # Fast literal object mapping replaces slow Add-Member construction loops
            $ReportCollection.Add([PSCustomObject]@{
                "Display Name"        = $Mailbox.DisplayName
                "Email Address"       = $PrimarySmtp
                "Sync Age (Days)"     = $SyncDays
                "DeviceID"            = $Device.DeviceID
                "DeviceAccessState"   = $Device.DeviceAccessState
                "DeviceModel"         = $Device.DeviceModel
                "DeviceType"          = $Device.DeviceType
                "DeviceFriendlyName"  = $Device.DeviceFriendlyName
                "DeviceOS"            = $Device.DeviceOS
                "LastSyncAttemptTime" = $Device.LastSyncAttemptTime
                "LastSuccessSync"     = $Device.LastSuccessSync
            })
        }
    }
    $CurrentIndex++
}

Write-Progress -Activity "Processing..." -Completed
if ($ReportCollection.Count -gt 0) {
    $ReportCollection | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nAnalysis complete! Mobile device matrix compiled to: $ExportPath" -ForegroundColor Green
} else {
    Write-Host "`nNo mobile devices matched your specified filter criteria." -ForegroundColor Yellow
}
