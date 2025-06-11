# Define source and destination Exchange servers
$Source = Read-Host "Enter source FQDN (e.g. exchange.domain.com)"
$Target = Read-Host "Enter target FQDN (e.g. exchange2.domain.com)"

# Get all Receive Connectors from the source server
$connectors = Get-ReceiveConnector -Server $Source | Where {$_.Identity -Notlike "*Default*" -and $_.Identity -NotLike "*Proxy*" -and $_.Identity -Notlike "*FrontEnd*"}

foreach ($connector in $connectors) {
    $Name = $connector.Name
    $bindings = $connector.Bindings
    $remoteIPRanges = $connector.RemoteIPRanges
    $fqdn = $connector.Fqdn
    $authMechanism = $connector.AuthMechanism
    $permissionGroups = $connector.PermissionGroups
    $maxMessageSize = $connector.MaxMessageSize
    $transportRole = $connector.TransportRole
    $tlsDomainCapabilities = $connector.TlsDomainCapabilities


Try {
    New-ReceiveConnector -Name $name `
                         -Server $Target `
                         -Bindings $bindings `
                         -RemoteIPRanges $remoteIPRanges `
                         -Fqdn $fqdn `
                         -AuthMechanism $authMechanism `
                         -ProtocolLoggingLevel Verbose `
                         -MaxMessageSize $maxMessageSize `
                         -TransportRole $transportRole `
                         -TlsDomainCapabilities $tlsDomainCapabilities
                         
                         
    }
Catch
{
Write-Host "Failed to make $Name"
}

    }
