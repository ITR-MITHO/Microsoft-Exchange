<#
.SYNOPSIS
    Optimizes IIS log configuration and purges aged Exchange/IIS log files.
.DESCRIPTION
    Updates IIS log rollover to Hourly for faster lock release and purges 
    IIS and Exchange log files older than 10 days using fast engine filtering.
.OUTPUTS
    Console progress bars and completion summary.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}

# 1. Enforce Module Prerequisites & Fast IIS Configuration Check
if (Get-Module -ListAvailable -Name WebAdministration) {
    Import-Module WebAdministration -ErrorAction SilentlyContinue
    
    # Standardize log generation period to Hourly to prevent massive file locks
    $CurrentPeriod = Get-WebConfigurationProperty -Filter /system.applicationHost/sites/siteDefaults/logFile -Name Period
    if ($CurrentPeriod.Value -ne "Hourly") {
        Set-WebConfigurationProperty -Filter /system.applicationHost/sites/siteDefaults/logFile -Name "period" -value "Hourly"
        Write-Host "IIS Log Rollover configuration updated to: Hourly" -ForegroundColor Yellow
    }
}

$ThresholdDate = (Get-Date).AddDays(-10)
$LogDirectories = [System.Collections.Generic.List[string]]::new()

# 2. Dynamically Identify All Active IIS Log Locations
if (Test-Path "IIS:\Sites") {
    Get-ChildItem "IIS:\Sites" | ForEach-Object {
        $Path = $_.logFile.directory
        if ($Path -match '%SystemDrive%') {
            $Path = $Path -replace '%SystemDrive%', $env:SystemDrive
        }
        # Only add the path if it exists and isn't already in our list
        if (-not [string]::IsNullOrEmpty($Path) -and (Test-Path $Path) -and (-not $LogDirectories.Contains($Path))) {
            [void]$LogDirectories.Add($Path)
        }
    }
}

# Fallback to standard default route ONLY if no directories were found via the IIS provider
if ($LogDirectories.Count -eq 0) {
    $DefaultIisPath = Join-Path $env:SystemDrive "Inetpub\Logs\LogFiles"
    if (Test-Path $DefaultIisPath) { 
        [void]$LogDirectories.Add($DefaultIisPath) 
    }
}

# Append the primary Exchange Logging directory root (ensuring no duplicates here either)
if ($env:ExchangeInstallPath) {
    $ExchangeLogPath = Join-Path $env:ExchangeInstallPath "Logging"
    if (-not $LogDirectories.Contains($ExchangeLogPath)) {
        [void]$LogDirectories.Add($ExchangeLogPath)
    }
}

# 3. High-Speed File Purge Engine
Write-Host "Starting log purge operations (Items older than: $($ThresholdDate.ToString('yyyy-MM-dd')))..." -ForegroundColor Cyan

foreach ($TargetDir in $LogDirectories) {
    if (-not (Test-Path $TargetDir)) { continue }
    Write-Host "Scanning directory: $TargetDir" -ForegroundColor DarkCyan
    
    # Optimization: Use -Filter directly inside Get-ChildItem to do the filtering at the OS level.
    # This is up to 4x faster than passing everything to Where-Object.
    $LogFiles = Get-ChildItem -Path $TargetDir -Recurse -Filter "*.log" -File -ErrorAction SilentlyContinue

    # Use native .Where() extension method instead of pipeline filtering (much faster in memory)
    $ExpiredFiles = $LogFiles.Where({ $_.LastWriteTime -lt $ThresholdDate })

    if ($ExpiredFiles.Count -gt 0) {
        Write-Host "Deleting $($ExpiredFiles.Count) expired log files..." -ForegroundColor Yellow
        foreach ($File in $ExpiredFiles) {
            try {
                Remove-Item -Path $File.FullName -Force -ErrorAction Stop
            } catch {
                # Handle gracefully if the log file is actively locked by Exchange/IIS worker processes
                continue 
            }
        }
    }
}

Write-Host "Log maintenance operation complete." -ForegroundColor Green
