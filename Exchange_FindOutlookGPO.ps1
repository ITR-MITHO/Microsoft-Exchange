<#
.SYNOPSIS
    Searches all Domain Group Policies for a specific keyword.
.DESCRIPTION
    Queries a local Domain Controller and efficiently parses GPO report data 
    to identify policies containing a specified string (e.g., 'Outlook').
.OUTPUTS
    Outputs a list of matching GPO names directly to the console host.
#>

Import-Module ActiveDirectory -ErrorAction SilentlyContinue

Write-Host "Locating an active Domain Controller..." -ForegroundColor Cyan
$DC = (Get-ADDomainController -Discover).HostName

Write-Host "Searching for Group Policies on $DC..." -ForegroundColor Yellow

# Define the search term cleanly outside the block
$SearchTerm = "Outlook"

$MatchedGPOs = Invoke-Command -ComputerName $DC -ArgumentList $SearchTerm {
    param($Keyword)
    
    Import-Module GroupPolicy -ErrorAction SilentlyContinue
    $AllGPOs = Get-GPO -All
    $Results = [System.Collections.Generic.List[string]]::new()

    foreach ($GPO in $AllGPOs) {
        try {
            $GpoXml = Get-GPOReport -Guid $GPO.Id -ReportType XML -ErrorAction Stop
            if ($GpoXml -like "*$Keyword*") {
                $Results.Add($GPO.DisplayName)
            }
        } catch {
            Write-Warning "Failed to generate report for GPO: $($GPO.DisplayName)"
        }
    }
    return $Results
    }

if ($MatchedGPOs) {
    Write-Host "`n=== Matching GPOs Found ===" -ForegroundColor Green
    foreach ($GpoName in $MatchedGPOs) {
        Write-Host " -> $GpoName" -ForegroundColor Green
    }
} else {
    Write-Host "No GPOs matched the search term: $SearchTerm" -ForegroundColor DarkYellow
}
