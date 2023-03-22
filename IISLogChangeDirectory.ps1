$Location = Read-Host "Enter the new location. (e.g. F:\Logs\)"

# Crate new IIS-directories
New-Item -ItemType Directory -Path $Location
New-Item -ItemType Directory -Path $Location

# Sets IIS to log in the new directories  
Import-Module WebAdministration
Set-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory  -value "$Location"
Set-ItemProperty "IIS:\Sites\Exchange Back End" -name logFile.directory  -value "$Location"
