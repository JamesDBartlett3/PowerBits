
# Title: Match-DatasetsWithWorkspaces
# Author: @JamesDBartlett3
# Parameters: $ds (list of dataset IDs)
# Returns: Table with two columns: DatasetId and WorkspaceId
# Usage: Match-DatasetsWithWorkspaces $ds

Param(
  $ds
)

$hadToLogin = $false
$ignoreWorkspaces = "Azure DevOps Dashboard", "Microsoft Project Web App", "Power BI Premium Capacity Metrics"
$obj = @{}

try {
  Get-PowerBIAccessToken | Out-Null
} catch {
  $hadToLogin = $true
  Login-PowerBIServiceAccount | Out-Null
} finally {
  $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
    Where-Object {$_.Type -EQ "Workspace" -AND $_.Name -NotIn $ignoreWorkspaces} |
    Select-Object Id
  $datasets = $null
  ForEach($w in $workspaces) {
    $workspaceId = $w.Id
    $datasets = Get-PowerBIDataset -WorkspaceId $workspaceId -ErrorAction "SilentlyContinue" |
      Select-Object -Property Id |
      Where-Object -Property Id -In $ds.Id
    ForEach($d in $datasets) {
      $obj.Add($d.Id, $workspaceId)
    }
  }
}

if($hadToLogin) {
  Logout-PowerBIServiceAccount
}

return $obj