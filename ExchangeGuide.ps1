# Microsoft Exchange Server INSTALLATION
# Install Exchange Server Prerequisites from here: https://learn.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites
# Download the newest Exchange version here: https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /PrepareSchema
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /PrepareAD /OrganizationName:"ITR-MITHO"
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /Mode:Install /Roles:Mailbox /on:"ITR-MITHO"


# Microsoft Exchange Server Cumulative UPGRADE
# Download the newest Exchange version here: https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /PrepareSchema
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /PrepareAD /OrganizationName:"ITR-MITHO"
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /Mode:Upgrade


# Verify that Exchange is healthy after install or upgrade:
cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/HealthCheck.ps1)"
