$url = "https://download.microsoft.com/download/ff69247e-f90f-495d-a846-0de5527f3849/ExchangeSubscriptionEdition-KB5128895-x64-en.exe"
$destinationFolder = "C:\Temp"
$fileName = "ExchangeSubscriptionEdition-KB5128895-x64-en.exe"
$destinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName

if (-not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
}

$ProgressPreference = 'SilentlyContinue'
Write-Host "Downloading $fileName to $destinationFolder..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $url -OutFile $destinationPath -UseBasicParsing

if (Test-Path -Path $destinationPath) {
    Write-Host "Download complete. Starting installer with elevated privileges..." -ForegroundColor Green    
    Start-Process -FilePath $destinationPath -ArgumentList "/passive /forcerestart" -Verb RunAs -Wait
} else {
    Write-Error "Failed to locate downloaded file at $destinationPath"
}
