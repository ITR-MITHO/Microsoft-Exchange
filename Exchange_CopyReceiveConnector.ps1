<#
The script is designed to copy Receive connector settings from one server to another. 
Edited/Fixed by SUMOL

#>

# Define source and destination Exchange servers
$Source = Read-Host "Enter source server FQDN (e.g. exchange.domain.com)"
$Target = Read-Host "Enter target server FQDN (e.g. exchange2.domain.com)"
 
# Ask for IPs to handle explicit bindings (e.g., swapping 192.168.1.50:25 to 192.168.1.51:25)
$SourceIP = Read-Host "Enter source server IPv4 (Leave blank if using 0.0.0.0)"
$TargetIP = Read-Host "Enter target server IPv4 (Leave blank if using 0.0.0.0)"
 
# Get all custom Receive Connectors
$connectors = Get-ReceiveConnector -Server $Source | Where-Object {
    $_.Identity -notmatch "Default|Proxy|FrontEnd"
}
 
foreach ($connector in $connectors) {
    Write-Host "`nProcessing connector: $($connector.Name)" -ForegroundColor Cyan
    # Handle IP Bindings via string replacement to avoid read-only object errors
    $newBindings = $connector.Bindings | ForEach-Object {
        $bindingStr = $_.ToString()
        if (![string]::IsNullOrWhiteSpace($SourceIP) -and ![string]::IsNullOrWhiteSpace($TargetIP)) {
            $bindingStr = $bindingStr -replace $SourceIP, $TargetIP
        }
        $bindingStr
    }

    #Since you can't "copy" Custom permissiongroups, remove Custom and apply the permission afterwards
    $PermGroups = $connector.PermissionGroups
    $PermGroups = $PermGroups -replace ", Custom", ""

    #If the old Fqdn is equal to the old servername, then change the new Fqdn to the new servername
    $newFQDN = $connector.Fqdn
    if ($connector.Fqdn -like $Source)
    {
        Write-Host "Old connector FQDN:" $connector.Fqdn
        Write-Host "New connector FQDN:" $Target
        $newFQDN = $Target
    }

 
#Build the parameter hashtable (Splatting)
    $connectorParams = @{
        Name                        = $connector.Name
        Server                      = $Target
        Bindings                    = $newBindings
        RemoteIPRanges              = $connector.RemoteIPRanges
        Fqdn                        = $newFQDN
        AuthMechanism               = $connector.AuthMechanism
        PermissionGroups            = $PermGroups
        ProtocolLoggingLevel        = 'Verbose'
        MaxMessageSize              = $connector.MaxMessageSize
        TransportRole               = $connector.TransportRole
        TlsDomainCapabilities       = $connector.TlsDomainCapabilities
        MessageRateLimit            = $connector.MessageRateLimit
        MaxRecipientsPerMessage     = $connector.MaxRecipientsPerMessage
        ConnectionInactivityTimeout = $connector.ConnectionInactivityTimeout
        ConnectionTimeout           = $connector.ConnectionTimeout
        RequireEHLODomain           = $connector.RequireEHLODomain
        ErrorAction                 = 'Stop'
    }
 
    # Exchange throws an error if you pass a null or empty Banner, so we only add it if it exists
    if (![string]::IsNullOrWhiteSpace($connector.Banner)) {
        $connectorParams.Add('Banner', $connector.Banner)
    }
 
    Try {        
        $newConnector = New-ReceiveConnector @connectorParams
        Write-Host "Successfully created $($connector.Name) on $Target." -ForegroundColor Green
 
        $anonRelay = Get-ADPermission -Identity $connector.Identity | Where-Object {
            $_.User -like "*ANONYMOUS LOGON*" -and $_.ExtendedRights -like "ms-Exch-SMTP-Accept-Any-Recipient"
        }        

        if ($anonRelay) {
            Write-Host "Migrating Anonymous Relay extended rights to $($connector.Name)..." -ForegroundColor Yellow
            Add-ADPermission -Identity $newConnector.Identity `
                             -User "NT AUTHORITY\ANONYMOUS LOGON" `
                             -ExtendedRights "ms-Exch-SMTP-Accept-Any-Recipient" `
                             -ErrorAction Stop | Out-Null
            Write-Host "Permissions applied successfully." -ForegroundColor Green
        }
    }
    Catch {
        Write-Host "Failed to create or configure $($connector.Name)." -ForegroundColor Red
        Write-Host "Error Details: $_" -ForegroundColor DarkRed
    }
    Write-Host "NOTE! It's not checked if the copied receive-connectors should be disabled or enabled, so manually disable any which need to."
}
