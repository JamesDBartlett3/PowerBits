<#
  .SYNOPSIS
    Retrieves the status of all nodes in all Data Gateway clusters to which the user has access.
  
  .DESCRIPTION
    This script will retrieve the status of all nodes in all Data Gateway clusters to which you have access. 
    It will prompt you to authenticate with Microsoft Entra ID (f.k.a. Azure Active Directory) if you haven't already done so.
  
  .EXAMPLE
    .\Get-DataGatewayStatus.ps1
  
  .NOTES
    This script does NOT require Azure AD app registration, service principal creation, or any other special setup.
    The only requirements are:
    - The user must be able to run PowerShell (and install the DataGateway module, if it's not already installed).
    - The user must have permissions to query the Data Gateway service.

    TODO
      - Replace DataGateway module dependency with Invoke-RestMethod calls to the [GatewayClusters API](https://api.powerbi.com/v2.0/myorg/gatewayclusters).
      
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
  
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits/blob/main/PowerShell/Scripts/Get-DataGatewayStatus.ps1)
  
  .LINK
    [The author's blog](https://datavolume.xyz)
    
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
#>

#Requires -Modules DataGateway

begin {
  try {
    Get-DataGatewayAccessToken | Out-Null
  }
  catch {
    Write-Host '🔒 DataGatewayAccessToken required. Launching Microsoft Entra ID (f.k.a. Azure Active Directory) authentication dialog...'
    Start-Sleep -s 1
    Login-DataGatewayServiceAccount -WarningAction SilentlyContinue | Out-Null
  }
  Write-Host '🔑 Power BI Access Token acquired.'
  Write-Host '⏳ Retrieving status of all accesssible Data Gateway nodes...'
}

process {
  Get-DataGatewayCluster | ForEach-Object {
    $clusterName = $_.Name
    $clusterId = $_.Id
    $_ | Select-Object -ExpandProperty MemberGateways | Select-Object -Property `
    @{l = 'ClusterId'; e = { $clusterId } }, 
    @{l = 'ClusterName'; e = { $clusterName } }, 
    @{l = 'NodeId'; e = { $_.Id } }, 
    @{l = 'NodeName'; e = { $_.Name } }, 
    @{l = 'ServerName'; e = { ($_.Annotation | ConvertFrom-Json).gatewayMachine } }, 
    Status, Version, VersionStatus, State
  }
}