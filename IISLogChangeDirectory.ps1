# Crate new IIS-directories
New-Item -ItemType Directory -Path F:\inetpub\logs\LogFiles\W3SVC1
New-Item -ItemType Directory -Path F:\inetpub\logs\LogFiles\W3SVC2

# Sets IIS to log in the new directories  
Import-Module WebAdministration
Set-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory  -value "F:\Inetpub\logs\LogFiles"
Set-ItemProperty "IIS:\Sites\Exchange Back End" -name logFile.directory  -value "F:\Inetpub\logs\LogFiles"
