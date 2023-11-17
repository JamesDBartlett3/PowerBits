#Requires -PSEdition Core
Function Export-PowerBIWorkspacesSecurity {
  #Requires -Modules MicrosoftPowerBIMgmt, ImportExcel
  try {
    Get-PowerBIAccessToken | Out-Null
  } catch {
    Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
  }
  finally {
    Write-Host '🔑 Power BI Access Token acquired.'
    $currentDate = Get-Date -UFormat "%Y-%m-%d_%H%M"
    $OutputFileName = "Power BI Workspace Security Audit ($currentDate).xlsx"
    # Get names of Workspaces to ignore from IgnoreList.json file
    # Most of these are template apps and/or auto-generated by Microsoft
    [PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "../IgnoreList.json") | ConvertFrom-Json
    [array]$ignoreWorkspaces = $ignoreObjects.IgnoreWorkspaces
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
      Where-Object {
        $_.State -NE "Deleted" -AND 
        $_.Type -EQ "Workspace" -AND 
        $_.IsOrphaned -EQ $False -AND 
        $_.Name -NotIn $ignoreWorkspaces -AND
        $_.Name -NotLike ".*"
      } | Select-Object -Property Id, Name |
      Sort-Object -Property Name -Unique
    $result = @()
    ForEach($w in $workspaces) {
      $workspaceName = $w.Name
      $workspaceId = $w.Id
      "Getting results for workspace: `e[38;2;255;0;0m$workspaceName`e[0m (Id: `e[38;2;0;255;0m$workspaceId`e[0m)" |
        Write-Host
      $pbiURL = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/users"
      $resultJson = Invoke-PowerBIRestMethod -Url $pbiURL -Method GET -ErrorAction SilentlyContinue
      $resultObject = ConvertFrom-Json -InputObject $resultJson 
      $result += $resultObject.Value |
        Select-Object @{n='workspaceId';e={$workspaceId}},
        @{n='workspaceName';e={$workspaceName}},
        @{n='userName';e={$_.displayName}},
        @{n='userRole';e={$_.groupUserAccessRight}},
        @{n='userType';e={$_.principalType}},
        @{n='emailAddress';e={$_.emailAddress}},
        @{n='identifier';e={$_.identifier}} |
        Sort-Object userRole, userName
      # Write-Host "Waiting 36 seconds to avoid hitting the API limit (200 req/hr)..."
      # Start-Sleep 36
    }
    $params = @{
      Path = Join-Path -Path $env:TEMP -ChildPath $OutputFileName
      Show = $true
      ClearSheet = $true
      AutoFilter = $true
      AutoSize = $true
      FreezeTopRow = $true
      BoldTopRow = $true
    }
    $result | 
      Select-Object -Property workspaceId, workspaceName, emailAddress, userRole, userType |
      Sort-Object -Property workspaceName, userRole, emailAddress | Export-Excel @params 
  }
}
Function Get-DataGatewayNodesStatus {
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
  #Requires -Modules DataGateway
  Write-Host '⏳ Retrieving status of all accesssible Data Gateway nodes...'
  try {
    Get-DataGatewayAccessToken | Out-Null
  } catch {
    Write-Host '🔒 DataGatewayAccessToken required. Launching Azure Active Directory authentication dialog...'
    Start-Sleep -s 1
    Login-DataGatewayServiceAccount -WarningAction SilentlyContinue | Out-Null
  } finally {
    Write-Host '🔑 Power BI Access Token acquired.'
    Get-DataGatewayCluster | ForEach-Object {
      $clusterName = $_.Name
      $clusterId = $_.Id
      $_ | Select-Object -ExpandProperty MemberGateways | Select-Object -Property `
      @{l = 'ClusterId'; e = { $clusterId }}, 
      @{l = 'ClusterName'; e = { $clusterName }}, 
      @{l = 'NodeId'; e = { $_.Id }}, 
      @{l = 'NodeName'; e = { $_.Name }}, 
      @{l = 'GatewayMachine'; e = { ($_.Annotation | ConvertFrom-Json).gatewayMachine }}, 
      Status, Version, VersionStatus, State
    }
  }
}
Function Export-PowerBIScannerApiData {
  <#
  .SYNOPSIS
    Title: Export-PowerBIScannerApiData
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)
  .DESCRIPTION
    This script will get all available data from the Power BI Scanner API and write it to a JSON file.
  .PARAMETER OutFile
    The destination path for the JSON file. Defaults to the Downloads folder.
  .PARAMETER OpenFile
    Specify to open the JSON file in the default application after it's created.
  .EXAMPLE
    .\Export-PowerBIScannerApiData.ps1 -OutFile "~\Downloads\PowerBIScannerApiData.json"
  .LINK
    https://github.com/JamesDBartlett3/PowerBits
  .LINK
    https://techhub.social/@JamesDBartlett3
  .LINK
    https://datavolume.xyz
  .NOTES
    - Requires the Power BI Management module: https://docs.microsoft.com/en-us/powershell/power-bi/overview?view=powerbi-ps
    - Requires the Power BI Scanner API to be enabled: https://learn.microsoft.com/en-us/power-bi/enterprise/service-admin-metadata-scanning#enabling-enhanced-metadata-scanning
    - Currently only works with workspaces that have been modified in the last 30 days
    - Tenants with a lot of workspaces may not work properly due to API rate limits
#> 
  Param(
    [Parameter(Mandatory = $false)]
    [string]$OutFile = '~\Downloads\PowerBIScannerApiData.json',
    [Parameter(Mandatory = $false)]
    [switch]$OpenFile
  )
  $currentDate = Get-Date -UFormat "%Y-%m-%d_%H%M"
  $OutFile = $OutFile -replace 'PowerBIScannerApiData.json', "PowerBIScannerApiData_$currentDate.json"
  $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()
  try {
    $headers = Get-PowerBIAccessToken
  }
  catch {
    Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
    if ($headers) {
      Write-Host '🔑 Power BI Access Token acquired. Proceeding...'
    } else {
      Write-Host '❌ Power BI Access Token not acquired. Exiting...'
      Exit
    }
  }
  $getWorkspacesUrl = 'https://api.powerbi.com/v1.0/myorg/admin/workspaces/modified?excludePersonalWorkspaces=True'
  # Send a GET request to the API endpoint
  $workspaceList = Invoke-RestMethod -Uri $getWorkspacesUrl -Method Get -Headers $headers
  # Create an object to hold the workspace IDs
  $workspaceIdsObject = [PSCustomObject]@{
    workspaces = @()
  }
  # Add the workspace IDs to the object
  foreach ($w in $workspaceList) {
    $workspaceIdsObject.workspaces += $w.id
  }
  # Convert the object to JSON
  $jsonBody = $workspaceIdsObject | ConvertTo-Json
  $startScanUrl = 'https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True'
  # Send a POST request to the API endpoint
  $startScanResponse = Invoke-RestMethod -Uri $startScanUrl -Method Post -Headers $headers -Body $jsonBody -ContentType 'application/json'
  $scanId = $startScanResponse.id
  $checkScanUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanStatus/$scanId"
  $scanStatus = 'NotStarted'
  # Check the scan status every 5 seconds until it's complete
  while ($scanStatus -ne 'Succeeded') {
    $checkScanResponse = Invoke-RestMethod -Uri $checkScanUrl -Method Get -Headers $headers
    Start-Sleep -s 5
    $scanStatus = $checkScanResponse.status
    Write-Host "Scan status: $scanStatus"
  }
  Write-Host 'Scan complete. Getting data...'
  $completedScanId = $checkScanResponse.id
  $getDataUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$completedScanId"
  # Send a GET request to the API endpoint
  $getDataResponse = Invoke-RestMethod -Uri $getDataUrl -Method Get -Headers $headers
  Write-Host "Writing data to file: $OutFile"
  # Write the data to a file
  $getDataResponse | ConvertTo-Json -Depth 100 | Out-File -FilePath $OutFile
  # Open the file in the default application if user passed the -OpenFile switch
  if ($OpenFile) {
    Invoke-Item $OutFile
  }
}
