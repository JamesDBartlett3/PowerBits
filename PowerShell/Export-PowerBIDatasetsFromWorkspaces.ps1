<#

  .SYNOPSIS
    Function: Export-PowerBIDatasetsFromWorkspaces
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

  .DESCRIPTION
    Export Power BI datasets from multiple workspaces in parallel

  .EXAMPLE
    Export-PowerBIDatasetsFromWorkspaces -OutputFolder C:\Datasets -ExtractWithPbiTools -SkipExistingFiles -ThrottleLimit 10

  .PARAMETER OutputFolder
    The folder where the datasets will be saved. If the folder does
    not exist, it will be created.

  .PARAMETER ExtractWithPbiTools
    If specified, exported PBIX datasets will be extracted with 
    pbi-tools after they are exported. Requires pbi-tools to be
    installed. See: https://pbi.tools

  .PARAMETER SkipExistingFiles
    If specified, existing files will be skipped. If not specified,
    existing files will be overwritten.

  .PARAMETER ThrottleLimit
    The maximum number of datasets that will be exported in parallel.
    Defaults to 1.

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must be allowed to download report PBIX files
        (see: "Download reports" setting in the Power BI Admin Portal).
    
    TODO


#>

Function Export-PowerBIDatasetsFromWorkspaces {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory=$false)][string]$OutputFolder,
    [Parameter(Mandatory=$false)][switch]$ExtractWithPbiTools,
    [Parameter(Mandatory=$false)][switch]$SkipExistingFiles,
    [Parameter(Mandatory=$false)][int]$ThrottleLimit = 1
  )

  #Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

  [string]$currentDateTime = Get-Date -UFormat "%Y%m%d_%H%M%S"
  [string]$fallbackDir = Join-Path -Path $env:TEMP -ChildPath "PowerBIWorkspaces"
  
  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

  Function Convert-PbixToProj {
    Param(
      [Parameter(Mandatory = $true)][string]$PbixPath,
      [Parameter(Mandatory = $true)][string]$ShortPath
    )
    try {
      Invoke-Expression pbi-tools | Out-Null
    } catch {
      Write-Error "'pbi-tools' command not found. See: https://pbi.tools/tutorials/getting-started-cli.html"
      Write-Warning $Error[0]
    }
    finally{
      if (!$Error[0]) {
        $command = "pbi-tools extract -pbixPath ""$PbixPath"""
        Write-Debug "Running command: $command"
        Write-Output "📦 Extracting: $ShortPath"
        Invoke-Expression $command | Out-Null
      }
    }
  }

  $fn_PbixToProj = ${function:Convert-PbixToProj}.ToString()

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

    $datasets = @()
    $datasetProperties = "" | Select-Object Name, Id, WebUrl, IsRefreshable, WorkspaceName, WorkspaceId

    $workspaces = Get-PowerBIWorkspace -Scope Organization -All -ErrorAction SilentlyContinue | 
      Where-Object {
        $_.Type -eq "Workspace" -and
        $_.State -eq "Active" -and
        $_.Name -notIn $ignoreWorkspaces
      } | Select-Object Name, Id | Sort-Object -Property Name |
      Out-ConsoleGridView -Title "Select Workspaces to Export"

    # For each workspace, find datasets with no corresponding report and add them to the $datasets array
    $workspaces | ForEach-Object {

      Write-Output "📁 $workspaceName"

      # Declare loop variables
      $workspaceName = $_.Name
      $workspaceId = $_.Id

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
        # If no corresponding report is found, add the current dataset to the $datasets array
        if (!$report) {
          $datasetProperties.Name = $datasetName
          $datasetProperties.Id = $datasetId
          $datasetProperties.WebUrl = $datasetWebUrl
          $datasetProperties.IsRefreshable = $datasetIsRefreshable
          $datasetProperties.WorkspaceName = $datasetWorkspaceName
          $datasetProperties.WorkspaceId = $datasetWorkspaceId
          $datasets += $datasetProperties
        }

      }

    }

    $datasets

  }
  
}

Export-PowerBIDatasetsFromWorkspaces