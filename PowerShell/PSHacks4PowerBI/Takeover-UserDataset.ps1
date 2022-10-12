<#

  .SYNOPSIS
    Function: Takeover-UserDataset
    Author: @JamesDBartlett3 (James D. Bartlett III)

  .DESCRIPTION
    - Take over a Power BI dataset that is currently configured by another user

  .PARAMETERS
    - $DatasetWorkspaceTable (table with two columns: DatasetId and WorkspaceId) -- set to output from Match-DatasetsWithWorkspaces

  .RETURNS
    - Nothing, if everything goes right ;-)

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must have permissions to access the workspace(s)
        in the Power BI service.

  .EXAMPLE
    Takeover-UserDataset $DatasetWorkspaceTable

  .TODO
    - Write as function
    - Re-implement token logic
    - 

  .ACKNOWLEDGEMENTS
    -

#>


#Requires -Modules MicrosoftPowerBIMgmt

Function Takeover-UserDataset {
  [CmdletBinding()]
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

}