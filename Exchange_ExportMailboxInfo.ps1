<#
.SYNOPSIS
    Comprehensive organization mailbox configuration and size auditing.
.DESCRIPTION
    Aggregates mailbox sizing, archival limits, AD states, and connection properties 
    into a structured inventory report.
.OUTPUTS
    $Home\Desktop\MailboxExport.csv
#>

$CsvPath = Join-Path $home "Desktop\MailboxExport.csv"

# 1. Privileged Context Validation
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevated administrative shell required. Exiting."
    break
}

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

# 2. Bulk Extraction Phase
Write-Host "Gathering Active Directory user status mappings..." -ForegroundColor Cyan
# Pre-cache only the 'Enabled' property for all users into a fast hash table
$AdUserMap = Get-ADUser -Filter * -Properties Enabled | Group-Object SamAccountName -AsHashTable -AsString

Write-Host "Gathering target mailboxes..." -ForegroundColor Cyan
$Mailboxes = Get-Mailbox -ResultSize Unlimited
$MailboxCount = $Mailboxes.Count
$Count = 1

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

# 3. Processing Engine Loop
foreach ($Mailbox in $Mailboxes) {
    $Sam = $Mailbox.SamAccountName
    
    # Progress UI Controller
    $DisplayName = '{0} ({1})' -f $Mailbox.DisplayName, $Mailbox.Name
    $Activity    = 'Working... [{0}/{1}]' -f $Count, $MailboxCount
    $Status      = 'Querying metadata profiles: {0}' -f $DisplayName
    Write-Progress -Status $Status -Activity $Activity -PercentComplete (($Count / $MailboxCount) * 100)

    # Fast AD memory lookup instead of hitting the network loop
    $AdEnabled = if ($AdUserMap.ContainsKey($Sam)) { $AdUserMap[$Sam].Enabled } else { $false }

    # Isolate Primary Mailbox Statistics
    $PrimaryStats = Get-MailboxStatistics -Identity $Sam -ErrorAction SilentlyContinue
    $PrimarySize  = if ($PrimaryStats.TotalItemSize.Value) { $PrimaryStats.TotalItemSize.Value.ToMB() } else { 0 }
    $LastLogon    = if ($PrimaryStats.LastLogonTime) { $PrimaryStats.LastLogonTime.ToString("dd-MM-yyyy") } else { $null }
    
    # Isolate Archive Mailbox Statistics (Only execute if the mailbox actually has an archive enabled)
    $ArchiveSize = "No Archive"
    if ($Mailbox.ArchiveStatus -ne "None") {
        $ArchiveStats = Get-MailboxStatistics -Identity $Sam -Archive -ErrorAction SilentlyContinue
        if ($ArchiveStats.TotalItemSize.Value) {
            $ArchiveSize = $ArchiveStats.TotalItemSize.Value.ToMB()
        }
    }

    $Results.Add([PSCustomObject]@{
        Username    = $Sam
        Name        = $Mailbox.DisplayName
        Email       = $Mailbox.PrimarySmtpAddress.ToString()
        UPN         = $Mailbox.UserPrincipalName.ToString() # Pulled directly from Exchange object
        Type        = $Mailbox.RecipientTypeDetails
        SizeInMB    = $PrimarySize
        ArchiveInMB = $ArchiveSize
        Retention   = $Mailbox.RetentionPolicy
        Forward     = $Mailbox.ForwardingAddress
        DB          = $PrimaryStats.Database
        LastLogon   = $LastLogon
        ADEnabled   = $AdEnabled
    })
    
    $Count++
}

# Complete UI progress bar visibility handle
Write-Progress -Activity "Working..." -Completed

# 4. Clear Export Alignment
$Results | Select-Object Username, Name, Email, UPN, Type, SizeInMB, ArchiveInMB, Retention, Forward, DB, LastLogon, ADEnabled | 
    Export-Csv -Path $CsvPath -NoTypeInformation -Encoding Unicode

Write-Host "Find your exported data here: $CsvPath" -ForegroundColor Green
