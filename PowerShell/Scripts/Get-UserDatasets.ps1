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
    
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
  #>

#Requires -PSEdition Core
#Requires -Modules MicrosoftPowerBIMgmt

Param(
  [parameter(Mandatory = $true)][string]$UserEmail
)

# Get names of Reports to ignore from IgnoreList.json file
# Most of these are template apps and/or auto-generated by Microsoft
[PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "../IgnoreList.json") | ConvertFrom-Json
[array]$ignoreReports = $ignoreObjects.IgnoreReports

try {
  Get-PowerBIAccessToken | Out-Null
} catch {
  Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID (f.k.a. Azure Active Directory) authentication dialog...'
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