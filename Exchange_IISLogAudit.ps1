<#
.SYNOPSIS
    Parses active IIS logs for a specific user identity or client IP address.
.DESCRIPTION
    Runs from an elevated shell context. Scans IIS logs modified within a defined 
    hourly window and buckets matching entries across core Exchange protocol logs.
.OUTPUTS
    $Home\Desktop\ExchangeLogs\*.log - Segmented search logs.
#>

$Hours = 12
$TargetDate = (Get-Date).AddHours(-$Hours)
$OutputDirectory = Join-Path $home "Desktop\ExchangeLogs"

if (-not (Get-Module -ListAvailable -Name WebAdministration)) {
    Write-Error "The WebAdministration module is missing. Please run on an IIS web server host."
    break
}
Import-Module WebAdministration -ErrorAction SilentlyContinue
$User = Read-Host "Enter username or client IP-address"
if ([string]::IsNullOrWhiteSpace($User)) {
    Write-Error "A valid search criteria string is required."
    break
}
if (-not (Test-Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}
Write-Host "Locating IIS log paths..." -ForegroundColor Cyan
$LogPaths = [System.Collections.Generic.List[string]]::new()

if (Test-Path "IIS:\Sites") {
    Get-ChildItem "IIS:\Sites" | ForEach-Object {
        $Path = $_.logFile.directory
        if ($Path -match '%SystemDrive%') {
            $Path = $Path -replace '%SystemDrive%', $env:SystemDrive
        }
        $SpecificSitePath = Join-Path $Path "W3SVC$($_.ID)"
        if (Test-Path $SpecificSitePath) { [void]$LogPaths.Add($SpecificSitePath) }
    }
}

if ($LogPaths.Count -eq 0) {
    $DefaultPath = Join-Path $env:SystemDrive "Inetpub\Logs\LogFiles\W3SVC1"
    if (Test-Path $DefaultPath) { [void]$LogPaths.Add($DefaultPath) }
}

Write-Host "Scanning log files modified since $($TargetDate.ToString('yyyy-MM-dd HH:mm:ss')) for '$User'..." -ForegroundColor Yellow
$TargetFiles = foreach ($Folder in $LogPaths) {
    Get-ChildItem -Path $Folder -Filter "*.log" -File -ErrorAction SilentlyContinue | 
        Where-Object { $_.LastWriteTime -gt $TargetDate }
}
if ($TargetFiles.Count -eq 0) {
    Write-Host "No log files modified within the last $Hours hours found." -ForegroundColor DarkYellow
    break
}

$Endpoints = @('Autodiscover', 'EWS', 'MAPI', 'OAB', 'OWA', 'ECP', 'ActiveSync')
$Streams = @{}
foreach ($Ep in $Endpoints) {
    $LogFile = Join-Path $OutputDirectory "$Ep.log"
    $Streams[$Ep] = [System.IO.StreamWriter]::new($LogFile, $false, [System.Text.Encoding]::UTF8)
}

# Loop through each log file exactly ONCE
foreach ($File in $TargetFiles) {
    try {
        $Reader = [System.IO.StreamReader]::new($File.FullName)
        while (($Line = $Reader.ReadLine()) -ne $null) {
            if ($Line -notlike "*$User*") { continue }
            switch -regex ($Line) {
                '/Autodiscover/'               { $Streams['Autodiscover'].WriteLine($Line); continue }
                '/EWS/'                        { $Streams['EWS'].WriteLine($Line); continue }
                '/MAPI/'                       { $Streams['MAPI'].WriteLine($Line); continue }
                '/OAB/'                        { $Streams['OAB'].WriteLine($Line); continue }
                '/OWA/'                        { $Streams['OWA'].WriteLine($Line); continue }
                '/ECP/'                        { $Streams['ECP'].WriteLine($Line); continue }
                '/Microsoft-Server-ActiveSync/' { $Streams['ActiveSync'].WriteLine($Line); continue }
            }
        }
        $Reader.Close()
    } catch {
        Write-Warning "Skipped file due to lock validation constraint: $($File.Name)"
    }
}
foreach ($Ep in $Endpoints) {
    $Streams[$Ep].Flush()
    $Streams[$Ep].Close()
    $Streams[$Ep].Dispose()
}

Write-Host "`nAnalysis complete! Extracted files exported to: $OutputDirectory" -ForegroundColor Green
