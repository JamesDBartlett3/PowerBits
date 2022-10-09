<#

  .SYNOPSIS
    Function: Get-UserDatasets
    Author: @JamesDBartlett3 (James D. Bartlett III)

  .DESCRIPTION
    - Get a list of all Power BI datasets marked as "configured by" a given user

  .PARAMETERS
    - $userEmail (string) -- email address of the user

  .RETURNS
    - List of dataset IDs

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must have permissions to access the workspace(s)
        in the Power BI service.

  .EXAMPLE
    Get-UserDatasets user@domain.tld

  .TODO
    - Write as function
    - Re-implement token logic
    - 

  .ACKNOWLEDGEMENTS
    -

#>

#Requires -Modules MicrosoftPowerBIMgmt

Param(
  [string]$userEmail
)

$hadToLogin = $false
$ignoreReports = "Report Usage Metrics Report", "Dashboard Usage Metrics Report"

try {
  Get-PowerBIAccessToken | Out-Null
}
catch {
  $hadToLogin = $true
  Connect-PowerBIServiceAccount | Out-Null
}
finally{
  $result = Get-PowerBIDataset -Scope Organization |
    Where-Object -Property ConfiguredBy -eq $userEmail |
    Where-Object -Property Name -NotIn $ignoreReports |
    Select-Object -Property Id, Name
}

if($hadToLogin) {
  Disconnect-PowerBIServiceAccount
}

return $result
