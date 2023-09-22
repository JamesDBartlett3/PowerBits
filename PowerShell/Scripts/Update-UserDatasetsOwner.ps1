<#
  .SYNOPSIS
    Function: Update-UserDatasetsOwner
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

  .DESCRIPTION
    Take over a Power BI dataset that is currently configured by another user

  .PARAMETER DatasetWorkspaceTable
    Table with two columns: DatasetId and WorkspaceId -- set to output from Join-UserDatasetsWithWorkspaces

  .EXAMPLE
    Update-UserDatasetsOwner -DatasetWorkspaceTable $DatasetWorkspaceTable

  .OUTPUTS
    Nothing, if everything goes right ;-)

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must have permissions to access the workspace(s)
        in the Power BI service.

    TODO
      - Write as function
      - Re-implement token logic
      - Testing
#>

#Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt

Param(
  [parameter(Mandatory = $true, ValueFromPipeline = $true)]$DatasetWorkspaceTable
)
try {
  Get-PowerBIAccessToken | Out-Null
} catch {
  Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
  Start-Sleep -s 1
  Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
} finally {
  Write-Host '🔑 Power BI Access Token acquired.'
  ForEach ($key in $DatasetWorkspaceTable.Keys) {
    $workspaceId = $DatasetWorkspaceTable[$key]
    $datasetId = $key
    $uri = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/Default.TakeOver"
  
    # Try to transfer ownership of dataset to current user
    try { 
      Invoke-PowerBIRestMethod -Url $uri -Method Post -ErrorAction "SilentlyContinue" -WarningAction "SilentlyContinue"
      # Show error if we had a non-terminating error which catch won't catch
      if (-Not $?) {
        $errmsg = Resolve-PowerBIError -Last
        $errmsg.Message
      }
    }
    catch {
      $errmsg = Resolve-PowerBIError -Last
      $errmsg.Message
    }
  }
}