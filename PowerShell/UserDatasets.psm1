#Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt

Function Get-UserDatasets {
  <#
    .SYNOPSIS
      Function: Get-UserDatasets
      Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)
  
    .DESCRIPTION
      Get a list of all Power BI datasets marked as "configured by" a given user
  
    .PARAMETER UserEmail
      Email address of the user
  
    .EXAMPLE
      Get-UserDatasets user@domain.tld
  
    .OUTPUTS
      List of dataset IDs
  
    .NOTES
      This function does NOT require Azure AD app registration, 
      service principal creation, or any other special setup.
      The only requirements are:
        - The user must be able to run PowerShell (and install the
          MicrosoftPowerBIMgmt module, if it's not already installed).
        - The user must have permissions to access the workspace(s)
          in the Power BI service.
  
      TODO
        - Re-implement token logic
        - Testing
  #>
  Param(
    [parameter(Mandatory = $true)][string]$UserEmail
  )
  
  # Get names of Workspaces and Reports to ignore from IgnoreList.json file
  # Most of these are template apps and/or auto-generated by Microsoft
  [PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "IgnoreList.json") | ConvertFrom-Json
  [array]$ignoreReports = $ignoreObjects.IgnoreReports

  try {
    Get-PowerBIAccessToken | Out-Null
  } catch {
    Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
  } finally {
    Write-Host '🔑 Power BI Access Token acquired.'
    $result = Get-PowerBIDataset -Scope Organization |
      Where-Object -Property ConfiguredBy -eq $UserEmail |
      Where-Object -Property Name -NotIn $ignoreReports |
      Select-Object -Property Id, Name
  }
  return $result
}
Function Join-UserDatasetsWithWorkspaces {
  <#
    .SYNOPSIS
      Function: Join-UserDatasetsWithWorkspaces
      Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)
  
    .DESCRIPTION
      - Audit the security settings of Power BI Workspaces
  
    .PARAMETER DatasetList
      - list of dataset IDs -- set to output from Get-UserDatasets
  
    .OUTPUTS
      - Table with two columns: DatasetId and WorkspaceId
  
    .EXAMPLE
      Join-UserDatasetsWithWorkspaces $DatasetList
  
    .NOTES
      This function does NOT require Azure AD app registration, 
      service principal creation, or any other special setup.
      The only requirements are:
        - The user must be able to run PowerShell (and install the
          MicrosoftPowerBIMgmt module, if it's not already installed).
        - The user must have permissions to access the workspace(s)
          in the Power BI service.
  
      TODO
        - Re-implement token logic
  #>
  Param(
    [parameter(Mandatory = $true, ValueFromPipeline = $true)]$DatasetList
  )
  # Get names of Workspaces and Reports to ignore from IgnoreList.json file
  # Most of these are template apps and/or auto-generated by Microsoft
  [PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "IgnoreList.json") | ConvertFrom-Json
  [array]$ignoreWorkspaces = $ignoreObjects.IgnoreWorkspaces

  $obj = @{}
  try {
    Get-PowerBIAccessToken | Out-Null
  } catch {
    Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
  } finally {
    Write-Host '🔑 Power BI Access Token acquired.'
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
Function Update-UserDatasetsOwner {
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
}