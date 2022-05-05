<#############################################################/

Get-DataGatewayNodesStatus.ps1

This script will retrieve the status of all nodes in 
all Data Gateway clusters to which you have access.

Author: @JamesDBartlett3

TODO: Replace DataGateway module dependency with 
    Invoke-RestMethod calls to the GatewayClusters API.
    https://api.powerbi.com/v2.0/myorg/gatewayclusters

/#############################################################>

#Requires -Modules DataGateway

Write-Output "Retrieving status of all accesssible Data Gateway nodes..."

try {
    Get-DataGatewayAccessToken | Out-Null
} catch {
    Write-Output "DataGatewayAccessToken required. Launching Azure Active Directory authentication dialog..."
    Start-Sleep -s 2
    Login-DataGatewayServiceAccount -WarningAction SilentlyContinue | Out-Null
}
finally {
    Get-DataGatewayCluster | ForEach-Object {
    $clusterName = $_.Name
    $clusterId = $_.Id
    $_ | Select-Object -ExpandProperty MemberGateways | 
    Select-Object -Property @{
        l="ClusterId"; e={$clusterId}}, @{
        l="ClusterName"; e={$clusterName}}, @{
        l="NodeId"; e={$_.Id}}, @{
        l="NodeName"; e={$_.Name}}, @{
        l="GatewayMachine"; e={
            ($_.Annotation | ConvertFrom-Json).gatewayMachine
            }
        }, Status, Version, VersionStatus, State
    }
}
