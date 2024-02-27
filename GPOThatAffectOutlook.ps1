CLS
Write-Host "Searching for group policies..." -ForegroundColor Yellow

$DC = (Get-ADDomainController | Select Name -First 1).Name
Invoke-Command -ComputerName $DC {

$AllGPO = Get-GPO -All -Domain $env:SERDNSDOMAIN
[string[]] $MatchedGPOList = @()

ForEach ($GPO in $AllGPO) { 
    $Report = Get-GPOReport -Guid $GPO.Id -ReportType XML 
    if ($Report -match 'Outlook') { 
        Write-Host "$($GPO.DisplayName)" -ForeGroundColor "Green"
        $MatchedGPOList += "$($GPO.DisplayName)";
} 
  }
    }
