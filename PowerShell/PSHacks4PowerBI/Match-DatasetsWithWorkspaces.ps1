<#

  .SYNOPSIS
    Function: Match-DatasetsWithWorkspaces
    Author: @JamesDBartlett3 (James D. Bartlett III)

  .DESCRIPTION
    - Audit the security settings of Power BI Workspaces

  .PARAMETERS
    - $DatasetList (list of dataset IDs) -- set to output from Get-UserDatasets

  .RETURNS
    - Table with two columns: DatasetId and WorkspaceId

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must have permissions to access the workspace(s)
        in the Power BI service.

  .EXAMPLE
    Match-DatasetsWithWorkspaces $DatasetList

  .TODO
    - Write as function
    - Re-implement token logic
    - 

  .ACKNOWLEDGEMENTS
    -

#>

Function Match-DatasetsWithWorkspaces {
  #Requires -Modules MicrosoftPowerBIMgmt
  [CmdletBinding()]
  Param(
    [parameter(Mandatory = $true, ValueFromPipeline = $true)]$DatasetList
  )

  $ignoreWorkspaces = "Azure DevOps Dashboard", "Microsoft Project Web App", "Power BI Premium Capacity Metrics"
  $obj = @{}

  try {
    Get-PowerBIAccessToken | Out-Null
  }
  catch {
    Write-Output "Power BI Access Token required. Launching authentication dialog..."
    Connect-PowerBIServiceAccount | Out-Null
  }
  finally {
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
    Where-Object { $_.Type -EQ "Workspace" -AND $_.Name -NotIn $ignoreWorkspaces } |
    Select-Object Id
    $datasets = $null
    ForEach ($w in $workspaces) {
      $workspaceId = $w.Id
      $datasets = Get-PowerBIDataset -WorkspaceId $workspaceId -ErrorAction "SilentlyContinue" |
      Select-Object -Property Id |
      Where-Object -Property Id -In $DatasetList.Id
      ForEach ($d in $datasets) {
        $obj.Add($d.Id, $workspaceId)
      }
    }
  }

  return $obj

}