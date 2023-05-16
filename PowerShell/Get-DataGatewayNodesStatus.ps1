<#
	.SYNOPSIS
		Function: Get-DataGatewayNodesStatus
		Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

	.DESCRIPTION
		This function will retrieve the status of all nodes in 
		all Data Gateway clusters to which you have access.

	.EXAMPLE
		Get-DataGatewayNodesStatus

	.NOTES
		This function does NOT require Azure AD app registration, 
		service principal creation, or any other special setup.
		The only requirements are:
		- The user must be able to run PowerShell (and install the
			DataGateway module, if it's not already installed).
		- The user must have permissions to query the Data Gateway
			service.

		TODO
			- Replace DataGateway module dependency with 
				Invoke-RestMethod calls to the GatewayClusters API.
				https://api.powerbi.com/v2.0/myorg/gatewayclusters
#>

Function Get-DataGatewayNodesStatus {
	#Requires -Modules DataGateway
	Write-Output "Retrieving status of all accesssible Data Gateway nodes..."
	try {
		Get-DataGatewayAccessToken | Out-Null
	} catch {
		Write-Output "DataGatewayAccessToken required. Launching Azure Active Directory authentication dialog..."
		Start-Sleep -s 1
		Login-DataGatewayServiceAccount -WarningAction SilentlyContinue | Out-Null
	} finally {
		Get-DataGatewayCluster | ForEach-Object {
			$clusterName = $_.Name
			$clusterId = $_.Id
			$_ | Select-Object -ExpandProperty MemberGateways | 
			Select-Object -Property `
				@{l = "ClusterId"; e = { $clusterId }}, 
				@{l = "ClusterName"; e = { $clusterName }}, 
				@{l = "NodeId"; e = { $_.Id }}, 
				@{l = "NodeName"; e = { $_.Name }}, 
				@{l = "GatewayMachine"; e = { ($_.Annotation | ConvertFrom-Json).gatewayMachine }}, 
				Status, Version, VersionStatus, State
		}
	}
}