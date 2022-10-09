<#

  .SYNOPSIS
    Function: Copy-PowerBIReportContentToBlankPBIXFile
    Author: @JamesDBartlett3 (James D. Bartlett III)

  .DESCRIPTION
    - This script will copy the contents of a published Power BI 
      report into a new report published from a blank PBIX 
    - This solves the problem where a Power BI report originally 
      created in the web browser cannot be downloaded from the 
      Power BI service as a PBIX file.

  .PARAMETERS
    - sourceReportId: The ID of the report to copy from
    - sourceWorkspaceId: The ID of the workspace to copy from
    - targetReportId: The ID of the report to copy to
    - targetWorkspaceId (optional): The ID of the workspace to copy to
    
  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must be allowed to download report PBIX files
        (see: "Download reports" setting in the Power BI Admin Portal).
      - The user must have "Contributor" or higher permissions 
        on the source and target workspace(s).

  .TODO
    - Add usage, help, and examples.
    - Add response codes and error handling.
    - Add support for downloading a blank PBIX file from a
      GitHub repo and publishing it to the same workspace as 
      the source report, so the user doesn't have to create
      and publish their own.
    
  .EXAMPLE
    - 

  .ACKNOWLEDGEMENTS
    This PS function was inspired by a blog article written by 
    one of the top minds in the Power BI space, Mathias Thierbach.
    Check out his article here: https://bit.ly/37ofVou
    And if you're not already using his pbi-tools for Power BI
    version control, you should check it out: https://pbi.tools
    
#>

#Requires -Modules MicrosoftPowerBIMgmt.Profile

Function Copy-PowerBIReportContentToBlankPBIXFile {

  Param(
    [parameter(Mandatory = $true)][string]$sourceReportId,
    [parameter(Mandatory = $true)][string]$sourceWorkspaceId,
    [parameter(Mandatory = $true)][string]$targetReportId,
    [parameter(Mandatory = $false)][string]$targetWorkspaceId = $sourceWorkspaceId
    )

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

  try {

    $headers = Get-PowerBIAccessToken

  } catch {

    Write-Output "Power BI Access Token required. Launching authentication dialog..."
    Start-Sleep -s 2
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken

  }

  finally {

    $pbiApiBaseUri = "https://api.powerbi.com/v1.0/myorg"
    $endpointUri = "$pbiApiBaseUri/groups/$targetWorkspaceId/reports/$targetReportId/UpdateReportContent"
    $body = @"
      {
        "sourceReport": {
          "sourceReportId": "$sourceReportId",
          "sourceWorkspaceId": "$sourceWorkspaceId"
        },
        "sourceType": "ExistingReport"
      }
"@
    $headers.Add("Content-Type","application/json")
    $response = Invoke-RestMethod -Uri $endpointUri -Method POST -Headers $headers -Body $body
    Write-Output $response

  }

}
