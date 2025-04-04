# Source and Target Exchange Servers
$SourceServer = Read-Host "Enter Source Server"
$TargetServer = Read-Host "Enter Target Server"


$Connectors = Get-ReceiveConnector -Server $SourceServer

foreach ($Connector in $Connectors) {
    Write-Host "Processing connector: $($Connector.Name)"
    
    # Filter out system connectors, but include custom and default connectors
    if ($Connector.Name -match "Default|Client|Outbound") {
        
        # Gather all properties to clone
        $NewConnectorName = $Connector.Name
        $Bindings = $Connector.Bindings
        $RemoteIPRanges = $Connector.RemoteIPRanges
        $AuthMechanism = $Connector.AuthMechanism
        $PermissionGroups = $Connector.PermissionGroups
        $FQDN = $Connector.FQDN
        $MaxMessageSize = $Connector.MaxMessageSize
        $ProtocolLoggingLevel = $Connector.ProtocolLoggingLevel
        $Enabled = $Connector.Enabled
        $Comment = $Connector.Comment
        $ConnectionTimeout = $Connector.ConnectionTimeout
        $ConnectionInactivityTimeout = $Connector.ConnectionInactivityTimeout
        $MaxInboundConnection = $Connector.MaxInboundConnection
        $MaxInboundConnectionPerSource = $Connector.MaxInboundConnectionPerSource
        $MaxMessageRate = $Connector.MessageRateLimit

        # Ensure correct formatting for RemoteIPRanges (from Ali Tajran's article)
        $FormattedRemoteIPRanges = $RemoteIPRanges | ForEach-Object {
            if ($_ -match ":") {
                # Handle IPv6 format
                $_
            } else {
                # Handle IPv4 format (no special processing)
                $_
            }
        }

        # Overwrite existing connector on the target (if exists)
        $ExistingConnector = Get-ReceiveConnector -Server $TargetServer | Where-Object { $_.Name -eq $NewConnectorName }
        if ($ExistingConnector) {
            Write-Host "Overwriting existing connector: $NewConnectorName"
            Set-ReceiveConnector -Identity $ExistingConnector.Identity `
                                 -Bindings $Bindings `
                                 -RemoteIPRanges $FormattedRemoteIPRanges `
                                 -AuthMechanism $AuthMechanism `
                                 -PermissionGroups $PermissionGroups `
                                 -FQDN $FQDN `
                                 -MaxMessageSize $MaxMessageSize `
                                 -ProtocolLoggingLevel $ProtocolLoggingLevel `
                                 -Enabled $Enabled `
                                 -Comment $Comment `
                                 -ConnectionTimeout $ConnectionTimeout `
                                 -ConnectionInactivityTimeout $ConnectionInactivityTimeout `
                                 -MaxInboundConnection $MaxInboundConnection `
                                 -MaxInboundConnectionPerSource $MaxInboundConnectionPerSource `
                                 -MessageRateLimit $MaxMessageRate

            Write-Host "Connector $NewConnectorName successfully updated on $TargetServer!"
        }
        else {
            Write-Host "Creating new connector: $NewConnectorName"
            New-ReceiveConnector -Name $NewConnectorName `
                                 -Server $TargetServer `
                                 -Bindings $Bindings `
                                 -RemoteIPRanges $FormattedRemoteIPRanges `
                                 -AuthMechanism $AuthMechanism `
                                 -PermissionGroups $PermissionGroups `
                                 -FQDN $FQDN `
                                 -MaxMessageSize $MaxMessageSize `
                                 -ProtocolLoggingLevel $ProtocolLoggingLevel `
                                 -Enabled $Enabled `
                                 -Comment $Comment `
                                 -ConnectionTimeout $ConnectionTimeout `
                                 -ConnectionInactivityTimeout $ConnectionInactivityTimeout `
                                 -MaxInboundConnection $MaxInboundConnection `
                                 -MaxInboundConnectionPerSource $MaxInboundConnectionPerSource `
                                 -MessageRateLimit $MaxMessageRate

            Write-Host "Connector $NewConnectorName successfully created on $TargetServer!"
        }
    }
}
