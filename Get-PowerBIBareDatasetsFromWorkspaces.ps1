<#

  .SYNOPSIS
    Function: Get-PowerBIBareDatasetsFromWorkspaces
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

  .DESCRIPTION
    Export Power BI datasets from multiple workspaces in parallel

  .EXAMPLE
    Get-PowerBIBareDatasetsFromWorkspaces -ThrottleLimit 10


  .PARAMETER ThrottleLimit
    The maximum number of datasets that will be exported in parallel.
    Defaults to 1.

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
    
    TODO

#>

Function Get-PowerBIBareDatasetsFromWorkspaces {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory=$false)][int]$ThrottleLimit = 1
  )

  #Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

  try {
    $headers = Get-PowerBIAccessToken
  } catch {
    Write-Output "🔒 Power BI Access Token required. Launching authentication dialog..."
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
  } finally {

    Write-Output "🔑 Power BI Access Token acquired."

    # If debugging, display the access token
    Write-Debug "Headers: `n $($headers.Keys)`n $($headers.Values)"

    # Define names of workspaces and reports to ignore
    # Most of these are auto-generated stuff from Microsoft
    [array]$ignoreWorkspaces = @(
      "Gen2 Utilization Metrics"
      , "Azure DevOps Dashboard"
      , "Microsoft Project Web App"
      , "Office365 Usage Analytics"
      , "Power BI Premium Capacity Metrics"
      , "Microsoft 365 Usage Analytics"
      , "Dataflow Snapshots"
      , "Power BI Release Plan"
      , "Power BI JSON Theme Guide"
      , "Apps Catalog on Microsoft AppSource"
      , "COVID-19 Global Report"
      , "COVID-19 US Tracking Report"
      , "Custom Visuals Exploration Tool"
      , "Template Apps Exploration Tool"
      , "Microsoft Fabric Capacity Metrics"
    )
    [array]$ignoreReports = @(
      "Report Usage Metrics Report"
      , "Dashboard Usage Metrics Report"
    )

    # Declare $bareDatasets array as a concurrent (thread-safe) PSObject
    $bareDatasets = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()

    # Get list of workspaces
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All -ErrorAction SilentlyContinue | 
      Where-Object {
        $_.Type -eq "Workspace" -and
        $_.State -eq "Active" -and
        $_.Name -notIn $ignoreWorkspaces
      } | Select-Object Name, Id | Sort-Object -Property Name |
      Out-ConsoleGridView -Title "Select Workspaces to Export"

    # For each workspace, find datasets with no corresponding report and add them to the $bareDatasets array
    $workspaces | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {

      # Declare local variables
      $workspaceName = $_.Name
      $workspaceId = $_.Id
      $localDatasets = $using:bareDatasets
      $datasetProperties = "" | Select-Object Name, Id, WebUrl, IsRefreshable, WorkspaceName, WorkspaceId

      # Get datasets from the workspace
      $workspaceDatasets = Get-PowerBIDataset -Scope Organization -WorkspaceId $workspaceId -ErrorAction SilentlyContinue |
        Where-Object {
          $_.IsRefreshable -eq $true -and
          $_.Name -notIn $ignoreReports
          } | Select-Object Name, Id, WebUrl, IsRefreshable, @{
              Name="WorkspaceName"; Expression={$workspaceName}
            }, @{
              Name="WorkspaceId"; Expression={$workspaceId}
            } | Sort-Object -Property Name
      
      # Get reports from the workspace
      $workspaceReports = Get-PowerBIReport -Scope Organization -WorkspaceId $workspaceId -ErrorAction SilentlyContinue |
        Where-Object {
          $_.Name -notIn $ignoreReports -and
          $_.WebUrl -notlike "*/rdlreports/*"
          } | Select-Object Name, Id, WebUrl, ReportType, DatasetId, @{
              Name="WorkspaceName"; Expression={$workspaceName}
            }, @{
              Name="WorkspaceId"; Expression={$workspaceId}
            } | Sort-Object -Property Name

      # For each dataset, check for any corresponding reports with the same name
      $workspaceDatasets | ForEach-Object {
        $datasetName = $_.Name
        $datasetId = $_.Id
        $datasetWebUrl = $_.WebUrl
        $datasetIsRefreshable = $_.IsRefreshable
        $datasetWorkspaceName = $_.WorkspaceName
        $datasetWorkspaceId = $_.WorkspaceId
        $report = $workspaceReports | Where-Object {
          $_.Name -eq $datasetName
        }

        # If no corresponding report is found, add the current dataset to the $bareDatasets array
        if (!$report) {
          $datasetProperties.Name = $datasetName
          $datasetProperties.Id = $datasetId
          $datasetProperties.WebUrl = $datasetWebUrl
          $datasetProperties.IsRefreshable = $datasetIsRefreshable
          $datasetProperties.WorkspaceName = $datasetWorkspaceName
          $datasetProperties.WorkspaceId = $datasetWorkspaceId
          $localDatasets.Add($datasetProperties)
        }

      }

    }

    $bareDatasets | Select-Object -Unique -Property * | Format-Table -AutoSize

  }

}

Get-PowerBIBareDatasetsFromWorkspaces -ThrottleLimit 8