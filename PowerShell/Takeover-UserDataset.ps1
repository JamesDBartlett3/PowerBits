
# Title: Takeover-UserDataset
# Author: @JamesDBartlett3
# Parameters: $DatasetWorkspaceTable (table with two columns: DatasetId and WorkspaceId)
# Returns: Nothing, if everything goes right ;-)
# Requires: $DatasetWorkspaceTable variable, set to output from Match-DatasetsWithWorkspaces.ps1
# Usage: Takeover-UserDataset $DatasetWorkspaceTable

#Requires -Modules MicrosoftPowerBIMgmt.Profile

Param(
  $DatasetWorkspaceTable
)

$hadToLogin = $false

try {
  Get-PowerBIAccessToken | Out-Null
}
catch {
  $hadToLogin = $true
  Connect-PowerBIServiceAccount | Out-Null
}
finally {
  ForEach($key in $DatasetWorkspaceTable.Keys) {
    $workspaceId = $DatasetWorkspaceTable[$key]
    $datasetId = $key
    $uri = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/Default.TakeOver"
  
    # Try to transfer ownership of dataset to current user
    try { 
      Invoke-PowerBIRestMethod -Url $uri -Method Post #-ErrorAction "SilentlyContinue" -WarningAction "SilentlyContinue"

      # Show error if we had a non-terminating error which catch won't catch
      if (-Not $?) {
        $errmsg = Resolve-PowerBIError -Last
        $errmsg.Message
      }
    } catch {
      $errmsg = Resolve-PowerBIError -Last
      $errmsg.Message
    }
  }
}

if($hadToLogin) {
  Disconnect-PowerBIServiceAccount
}
