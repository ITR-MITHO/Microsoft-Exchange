<#
.SYNOPSIS
    Purges aged Exchange/IIS log files.
.DESCRIPTION
    IIS and Exchange log files older than 10 days.
.OUTPUTS
    Console progress bars and completion summary.
#>

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *EXC* -ErrorAction SilentlyContinue
}


$ThresholdDate = (Get-Date).AddDays(-10)
$LogDirectories = [System.Collections.Generic.List[string]]::new()

if (Test-Path "IIS:\Sites") {
    Get-ChildItem "IIS:\Sites" | ForEach-Object {
        $Path = $_.logFile.directory
        if ($Path -match '%SystemDrive%') {
            $Path = $Path -replace '%SystemDrive%', $env:SystemDrive
        }
        if (-not [string]::IsNullOrEmpty($Path) -and (Test-Path $Path) -and (-not $LogDirectories.Contains($Path))) {
            [void]$LogDirectories.Add($Path)
        }
    }
}
if ($LogDirectories.Count -eq 0) {
    $DefaultIisPath = Join-Path $env:SystemDrive "Inetpub\Logs\LogFiles"
    if (Test-Path $DefaultIisPath) { 
        [void]$LogDirectories.Add($DefaultIisPath) 
    }
}
if ($env:ExchangeInstallPath) {
    $ExchangeLogPath = Join-Path $env:ExchangeInstallPath "Logging"
    if (-not $LogDirectories.Contains($ExchangeLogPath)) {
        [void]$LogDirectories.Add($ExchangeLogPath)
    }
}
Write-Host "Starting log purge operations (Items older than: $($ThresholdDate.ToString('yyyy-MM-dd')))..." -ForegroundColor Cyan
foreach ($TargetDir in $LogDirectories) {
    if (-not (Test-Path $TargetDir)) { continue }
    Write-Host "Scanning directory: $TargetDir" -ForegroundColor DarkCyan
    $LogFiles = Get-ChildItem -Path $TargetDir -Recurse -Filter "*.log" -File -ErrorAction SilentlyContinue
    $ExpiredFiles = $LogFiles.Where({ $_.LastWriteTime -lt $ThresholdDate })

    if ($ExpiredFiles.Count -gt 0) {
        Write-Host "Deleting $($ExpiredFiles.Count) expired log files..." -ForegroundColor Yellow
        foreach ($File in $ExpiredFiles) {
            try {
                Remove-Item -Path $File.FullName -Force -ErrorAction Stop
            } catch {
                continue 
            }
        }
    }
}

Write-Host "Log maintenance operation complete." -ForegroundColor Green
