#Requires -PSEdition Core
Function Checkpoint-PowerBIWorkspaceSecurity {
  <#
  .SYNOPSIS
    Exports a list of all Power BI workspaces and their members to an Excel file.
  .DESCRIPTION
    This script exports a list of all Power BI workspaces and their members to an Excel file.
    It first authenticates with Power BI using an access token. If the access token is not available, it prompts the user to authenticate with Microsoft Entra ID.
    It then retrieves a list of all workspaces in the organization, excluding those that are deleted, not of type "Workspace", orphaned, or listed in the IgnoreList.json file.
    The resulting list of workspaces and their members is then exported to an Excel file with a timestamp in the filename. 
    This can be useful for auditing and security purposes.
	.PARAMETER OutputFile
		Specifies the path and filename of the Excel file to be created. If not specified, the file will be created in the user's TEMP directory with a timestamp in the filename.
	.PARAMETER OpenFile
		Switch to open the Excel file after it is created.
	.EXAMPLE
		# Export a list of all Power BI workspaces and their members to an Excel file
		# in the user's TEMP directory, then open the file.
		.\Checkpoint-PowerBIWorkspaceSecurity.ps1 -OpenFile
  .NOTES
    ACKNOWLEDGEMENTS:
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits/blob/main/PowerShell/Scripts/Checkpoint-PowerBIWorkspaceSecurity.ps1)
  .LINK
    [The author's blog](https://datavolume.xyz)
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
#>
  Param(
    [Parameter(Mandatory=$false)][string]$OutputFile,
    [Parameter(Mandatory=$false)][switch]$OpenFile
  )
  #Requires -Modules MicrosoftPowerBIMgmt, ImportExcel
  begin{
    $headers = [System.Collections.Generic.Dictionary[[String],[String]]]::New()
    try {
      $headers = Get-PowerBIAccessToken
    }
    catch {
      Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID authentication dialog...' -ForegroundColor DarkYellow
      Start-Sleep -s 1
      Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
      $headers = Get-PowerBIAccessToken
      if (!$headers) {
        Write-Host '❌ Power BI Access Token not acquired. Exiting...' -ForegroundColor Red
        Exit
      }
    }
    $currentDate = Get-Date -UFormat "%Y-%m-%d_%H-%M-%S"
    $OutputFile = if(!($OutputFile)) { 
      Join-Path -Path $env:TEMP -ChildPath "Power BI Workspace Security Audit ($currentDate).xlsx"
    } else {
      $OutputFile
    }
    Write-Host '🔑 Power BI Access Token acquired.' -ForegroundColor Green
  }
  process {
    # Get names of Workspaces to ignore from IgnoreList.json file
    # Most of these are template apps and/or auto-generated by Microsoft
    [PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "../IgnoreList.json") | ConvertFrom-Json
    [array]$ignoreWorkspaces = $ignoreObjects.IgnoreWorkspaces
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
      Where-Object {
        $_.State -NE "Deleted" -AND
        $_.Type -EQ "Workspace" -AND
        $_.IsOrphaned -EQ $False -AND
        $_.Name -NotIn $ignoreWorkspaces
      } | Select-Object -Property Id, Name | Sort-Object -Property Name -Unique
    $result = @()
    ForEach ($w in $workspaces) {
      $workspaceName = $w.Name
      $workspaceId = $w.Id
      "Getting results for workspace: `e[38;2;255;0;0m$workspaceName`e[0m (Id: `e[38;2;0;255;0m$workspaceId`e[0m)" | Write-Host
      $pbiURL = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/users"
      $resultJson = Invoke-PowerBIRestMethod -Url $pbiURL -Method GET -ErrorAction SilentlyContinue
      $resultObject = ConvertFrom-Json -InputObject $resultJson
      $result += $resultObject.Value |
        Select-Object @{n = 'workspaceId'; e = { $workspaceId } },
        @{n = 'workspaceName'; e = { $workspaceName } },
        @{n = 'userName'; e = { $_.displayName } },
        @{n = 'userRole'; e = { $_.groupUserAccessRight } },
        @{n = 'userType'; e = { $_.principalType } },
        @{n = 'emailAddress'; e = { $_.emailAddress } },
        @{n = 'identifier'; e = { $_.identifier } } |
        Sort-Object userRole, userName
    }
    $params = @{
      Path         = $OutputFile
      Show         = $OpenFile
      ClearSheet   = $true
      AutoFilter   = $true
      AutoSize     = $true
      FreezeTopRow = $true
      BoldTopRow   = $true
    }
    $result |
      Select-Object -Property workspaceId, workspaceName, emailAddress, userRole, userType |
      Sort-Object -Property workspaceName, userRole, emailAddress | Export-Excel @params
  }
}
Function Get-DataGatewayStatus {
  <#
  .SYNOPSIS
    Retrieves the status of all nodes in all Data Gateway clusters to which the user has access.
  .DESCRIPTION
    This script will retrieve the status of all nodes in all Data Gateway clusters to which you have access. 
    It will prompt you to authenticate with Microsoft Entra ID if you haven't already done so.
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
      Write-Host '🔒 Data Gateway Access Token required. Launching Microsoft Entra ID authentication dialog...' -ForegroundColor DarkYellow
      Start-Sleep -s 1
      Login-DataGatewayServiceAccount -WarningAction SilentlyContinue | Out-Null
    }
    Write-Host '🔑 Data Gateway Access Token acquired.' -ForegroundColor Green
    Write-Host '⏳ Retrieving status of all accesssible Data Gateway nodes...' -ForegroundColor Yellow
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
}
Function Export-PowerBIScannerApiData {
  <#
  .SYNOPSIS
    Title: Export-PowerBIScannerApiData
    Author: James D. Bartlett III
  .DESCRIPTION
    This script will get all available data from the Power BI Scanner API and write it to a JSON file.
  .INPUTS
    - Parameters are currently the only way to pass input to this script
    - Pipeline inputs are not yet supported
  .OUTPUTS
    - A .json file containing all available data from the Power BI Scanner API
    - Pipeline outputs are not yet supported
  .PARAMETER OutFile
    The destination path for the JSON file. Defaults to "~\Downloads\PowerBIScannerApiData_{timestamp}.json"
  .PARAMETER OpenFile
    Specify to open the JSON file in the default application after it's created.
  .EXAMPLE
    .\Export-PowerBIScannerApiData.ps1 -OutFile "C:\temp\MyPowerBIScannerApiData.json"
    # Export data to "C:\temp\MyPowerBIScannerApiData.json"
  .EXAMPLE
    .\Export-PowerBIScannerApiData.ps1 -OpenFile
    # Export data to the default location ("~\Downloads\PowerBIScannerApiData_{timestamp}.json")
    # and open the file in the system's default .json file handler/editor
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits)
  .LINK
    [The author's blog](https://datavolume.xyz)
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
  .NOTES
    - Requires the Power BI Management module: https://docs.microsoft.com/en-us/powershell/power-bi/overview?view=powerbi-ps
    - Requires the Power BI Scanner API to be enabled: https://learn.microsoft.com/en-us/power-bi/enterprise/service-admin-metadata-scanning#enabling-enhanced-metadata-scanning
    - Currently only works with workspaces that have been modified in the last 30 days
    - Tenants with a lot of workspaces may not work properly due to API rate limits
    TODO:
      - Add parameters for all available API options
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>
  Param(
    [Parameter(Mandatory = $false)]
    [string]$OutFile = "$HOME\Downloads\PowerBIScannerApiData_$(Get-Date -UFormat '%Y-%m-%d_%H%M').json",
    [Parameter(Mandatory = $false)]
    [switch]$OpenFile
  )
  $currentDate = Get-Date -UFormat "%Y-%m-%d_%H%M"
  $OutFile = $OutFile -replace 'PowerBIScannerApiData.json', "PowerBIScannerApiData_$currentDate.json"
  $baseUrl = 'https://api.powerbi.com/v1.0/myorg/admin/workspaces'
  $headers = [System.Collections.Generic.Dictionary[[String],[String]]]::New()
  try {
    $headers = Get-PowerBIAccessToken
  }
  catch {
    Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID authentication dialog...' -ForegroundColor DarkYellow
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
    if (!$headers) {
      Write-Host '❌ Power BI Access Token not acquired. Exiting...' -ForegroundColor Red
      Exit
    }
  }
  $modifiedWorkspacesUrl = "$baseUrl/modified?excludePersonalWorkspaces=True"
  # Send a GET request to the API endpoint
  $workspaceList = Invoke-RestMethod -Uri $modifiedWorkspacesUrl -Method Get -Headers $headers
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
  $getInfoUrl = "$baseUrl/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True"
  # Send a POST request to the API endpoint
  $startScanResponse = Invoke-RestMethod -Uri $getInfoUrl -Method Post -Headers $headers -Body $jsonBody -ContentType 'application/json'
  $scanId = $startScanResponse.id
  $scanStatusUrl = "$baseUrl/scanStatus/$scanId"
  $scanStatus = ''
  # Check the scan status every 5 seconds until it's complete
  while ($scanStatus -ne 'Succeeded') {
    $checkScanResponse = Invoke-RestMethod -Uri $scanStatusUrl -Method Get -Headers $headers
    Start-Sleep -s 5
    $scanStatus = $checkScanResponse.status
    Write-Host "Scan status: $scanStatus"
  }
  Write-Host 'Scan complete. Getting data...'
  $completedScanId = $checkScanResponse.id
  $scanResultUrl = "$baseUrl/scanResult/$completedScanId"
  # Send a GET request to the API endpoint
  $getDataResponse = Invoke-RestMethod -Uri $scanResultUrl -Method Get -Headers $headers
  Write-Host "Writing data to file: $OutFile"
  # Write the data to a file
  $getDataResponse | ConvertTo-Json -Depth 100 | Out-File -FilePath $OutFile
  # Open the file in the default application if user passed the -OpenFile switch
  if ($OpenFile) {
    Invoke-Item $OutFile
  }
}
