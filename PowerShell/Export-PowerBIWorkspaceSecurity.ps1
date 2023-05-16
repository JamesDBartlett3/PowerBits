Function Export-PowerBIWorkspaceSecurity {

  #Requires -PSEdition Core
  #Requires -Modules MicrosoftPowerBIMgmt, ImportExcel

  try {
    Get-PowerBIAccessToken | Out-Null
  }
  catch {
    Write-Output "Power BI Access Token required. Launching authentication dialog..."
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
  }
  finally {

    $currentDate = Get-Date -UFormat "%Y-%m-%d_%H%M"
    $OutputFileName = "Power BI Workspace Security Audit ($currentDate).xlsx"

    $ignoreWorkspaces = @(
      "Azure DevOps Dashboard",
      "Microsoft Project Web App",
      "Power BI Premium Capacity Metrics"
    )

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
        Write-Output
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
        # Write-Output "Waiting 36 seconds to avoid hitting the API limit (200 req/hr)..."
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
      Select-Object -Property workspaceId, workspaceName, userName, emailAddress, identifier, userRole, userType |
      Sort-Object -Property workspaceName, userRole, userName | Export-Excel @params 

  }

}