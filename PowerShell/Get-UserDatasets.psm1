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

Function Get-UserDatasets {
  #Requires -PSEdition Core
  #Requires -Modules MicrosoftPowerBIMgmt
  Param(
    [parameter(Mandatory = $true)][string]$UserEmail
  )
  $ignoreReports = "Report Usage Metrics Report", "Dashboard Usage Metrics Report"
  try {
    Get-PowerBIAccessToken | Out-Null
  } catch {
    Write-Output "🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog..."
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
  } finally {
    Write-Output "🔑 Power BI Access Token acquired."
    $result = Get-PowerBIDataset -Scope Organization |
      Where-Object -Property ConfiguredBy -eq $UserEmail |
      Where-Object -Property Name -NotIn $ignoreReports |
      Select-Object -Property Id, Name
  }
  return $result
}