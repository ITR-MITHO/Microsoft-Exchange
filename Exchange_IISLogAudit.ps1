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

# 1. Enforce Module Prerequisites & Target Extraction
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

# Ensure output repository path structure exists
if (-not (Test-Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

# 2. Dynamically Identify Active Web Site Log Directories
Write-Host "Locating IIS log paths..." -ForegroundColor Cyan
$LogPaths = [System.Collections.Generic.List[string]]::new()

if (Test-Path "IIS:\Sites") {
    Get-ChildItem "IIS:\Sites" | ForEach-Object {
        $Path = $_.logFile.directory
        if ($Path -match '%SystemDrive%') {
            $Path = $Path -replace '%SystemDrive%', $env:SystemDrive
        }
        # Append the site specific W3SVC folder sub-directory block
        $SpecificSitePath = Join-Path $Path "W3SVC$($_.ID)"
        if (Test-Path $SpecificSitePath) { [void]$LogPaths.Add($SpecificSitePath) }
    }
}

# Fallback pathing safety rule
if ($LogPaths.Count -eq 0) {
    $DefaultPath = Join-Path $env:SystemDrive "Inetpub\Logs\LogFiles\W3SVC1"
    if (Test-Path $DefaultPath) { [void]$LogPaths.Add($DefaultPath) }
}

# 3. High-Speed Single Pass Log Processing Engine
Write-Host "Scanning log files modified since $($TargetDate.ToString('yyyy-MM-dd HH:mm:ss')) for '$User'..." -ForegroundColor Yellow

# Identify target files quickly at the engine boundary without sorting
$TargetFiles = foreach ($Folder in $LogPaths) {
    Get-ChildItem -Path $Folder -Filter "*.log" -File -ErrorAction SilentlyContinue | 
        Where-Object { $_.LastWriteTime -gt $TargetDate }
}

if ($TargetFiles.Count -eq 0) {
    Write-Host "No log files modified within the last $Hours hours found." -ForegroundColor DarkYellow
    break
}

# Initialize StreamWriter memory engines for high-speed concurrent writing
$Endpoints = @('Autodiscover', 'EWS', 'MAPI', 'OAB', 'OWA', 'ECP', 'ActiveSync')
$Streams = @{}
foreach ($Ep in $Endpoints) {
    $LogFile = Join-Path $OutputDirectory "$Ep.log"
    # Overwrite old files, append mode active, uses standard encoding
    $Streams[$Ep] = [System.IO.StreamWriter]::new($LogFile, $false, [System.Text.Encoding]::UTF8)
}

# Loop through each log file exactly ONCE
foreach ($File in $TargetFiles) {
    try {
        # High speed system stream reader avoids filling RAM arrays
        $Reader = [System.IO.StreamReader]::new($File.FullName)
        while (($Line = $Reader.ReadLine()) -ne $null) {
            
            # Fast check: If the user string isn't in the line, skip immediately
            if ($Line -notlike "*$User*") { continue }

            # Match and sort the line to its respective stream bucket instantly
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

# Safely close and flush all open StreamWriters
foreach ($Ep in $Endpoints) {
    $Streams[$Ep].Flush()
    $Streams[$Ep].Close()
    $Streams[$Ep].Dispose()
}

Write-Host "`nAnalysis complete! Extracted files exported to: $OutputDirectory" -ForegroundColor Green
