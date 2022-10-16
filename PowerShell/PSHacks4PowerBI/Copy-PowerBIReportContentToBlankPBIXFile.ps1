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
    - Testing
    - Add usage, help, and examples.
    - Rename the function to something more accurate to its current capabilities.
    
  .EXAMPLE
    - 
  
  .ACKNOWLEDGEMENTS
    This PS function was inspired by a blog article written by 
    one of the top minds in the Power BI space, Mathias Thierbach.
    Check out his article here: https://bit.ly/37ofVou
    And if you're not already using his pbi-tools for Power BI
    version control, you should check it out: https://pbi.tools
  
#>

Function Copy-PowerBIReportContentToBlankPBIXFile {
  
  #Requires -PSEdition Core
  #Requires -Modules MicrosoftPowerBIMgmt
  
  [CmdletBinding()]
  Param(
    [parameter(Mandatory = $true)][string]$sourceReportId,
    [parameter(Mandatory = $true)][string]$sourceWorkspaceId,
    [parameter(Mandatory = $false)][string]$targetReportId,
    [parameter(Mandatory = $false)][string]$targetWorkspaceId = $sourceWorkspaceId,
    [Parameter(Mandatory = $false)][string]$blankPbix,
    [Parameter(Mandatory = $false)][string]$outFile
  )
  
  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  
  [string]$blankPbixTempFile = Join-Path -LiteralPath $env:TEMP -ChildPath "blank.pbix"
  [array]$validPbixContents = @("[Content_Types].xml", "Version", "Layout", "Metadata")
  
  [bool]$blankPbixIsUrl = $blankPbix.StartsWith("http")
  [bool]$localFileExists = Test-Path $blankPbix
  [bool]$remoteFileIsValid = $false
  [bool]$localFileIsValid = $false
  [bool]$defaultFileIsValid = $false
  
  Function FileIsBlankPbix($file) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
    $fileIsPbix = @($validPbixContents | Where-Object {$zip.Entries.Name -Contains $_}).Count -gt 0
    $fileIsBlank = (Get-Item $file).length / 1KB -lt 20
    $zip.Dispose()
    if($fileIsPbix -and $fileIsBlank) {
      Write-Debug "$file is a valid blank pbix file."
      return $true
      
    }
    else {
      Write-Error "$file is NOT a valid blank pbix file."
      return $false
    }
  }
  
  # If user did not specify a target report ID, use a blank PBIX file
  If(!$targetReportId) {
    # If user specified a URL to a file, download and validate it as a blank PBIX file
    if ($blankPbixIsUrl){
      Write-Debug "Downloading file: $blankPbix..."
      Invoke-WebRequest -Uri $blankPbix -OutFile $blankPbixTempFile
      Write-Debug "Validating downloaded file..."
      $remoteFileIsValid = FileIsBlankPbix($blankPbixTempFile)
    }
    
    # If user specified a local path to a file, validate it as a blank PBIX file
    elseif ($localFileExists) {
      Write-Debug "Validating user-supplied file: $blankPbix..."
      $localFileIsValid = FileIsBlankPbix($blankPbix)
    }
    
    # If user didn't specify a blank PBIX file, check for a valid blank PBIX in the temp location
    elseif (Test-Path $blankPbixTempFile) {
      Write-Debug "Validating pbix file found in temp location: $blankPbixTempFile..."
      $defaultFileIsValid = FileIsBlankPbix($blankPbixTempFile)
    }
    
    # If user did not specify a blank PBIX file, and a valid blank PBIX is not in the temp location,
    # download one from GitHub and check if it's valid and blank
    else {
      Write-Debug "Downloading a blank pbix file from GitHub to $blankPbixTempFile..."
      $blankPbixUri = "https://github.com/JamesDBartlett3/PowerBits/raw/main/Misc/blank.pbix"
      Invoke-WebRequest -Uri $blankPbixUri -OutFile $blankPbixTempFile
      $defaultFileIsValid = FileIsBlankPbix($blankPbixTempFile)
    }
    
    # If we downloaded a valid blank PBIX file, use it.
    if ($remoteFileIsValid -or $defaultFileIsValid) {
      $blankPbix = $blankPbixTempFile
    }
    
    # If a valid blank PBIX file could not be obtained by any of the above methods, throw an error.
    if (!$targetReportId -and !$localFileIsValid -and !$remoteFileIsValid -and !$defaultFileIsValid) {
      Write-Error "No targetReportId specified & no valid blank PBIX file found. Please specify one or the other."
      return
    }
    
    [bool]$pbixIsValid = ($localFileIsValid -or $remoteFileIsValid -or $defaultFileIsValid)
    
  }
  
  try {
    $headers = Get-PowerBIAccessToken
  }
  
  catch {
    Write-Output "Power BI Access Token required. Launching authentication dialog..."
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
  }
  
  finally {
    Write-Debug "Target Report ID is null: $(!$targetReportId)"
    $pbiApiBaseUri = "https://api.powerbi.com/v1.0/myorg"
    
    # If a valid blank PBIX was found, publish it to the target workspace
    if ($pbixIsValid) {
      Write-Debug "Publishing $blankPbix to target workspace..."
      $publishResponse = New-PowerBIReport -Path $blankPbix -WorkspaceId $targetWorkspaceId -ConflictAction CreateOrOverwrite
      Write-Debug "Response: $publishResponse"
      $targetReportId = $publishResponse.Id
    }
    
    # Assemble the UpdateReportContent API URI and request body
    $updateReportContentEndpoint = "$pbiApiBaseUri/groups/$targetWorkspaceId/reports/$targetReportId/UpdateReportContent"
    $body = @"
      {
        "sourceReport": {
          "sourceReportId": "$sourceReportId",
          "sourceWorkspaceId": "$sourceWorkspaceId"
        },
        "sourceType": "ExistingReport"
      }
"@
    # Update the target report with the source report's content
    $headers.Add("Content-Type", "application/json")
    $response = Invoke-RestMethod -Uri $updateReportContentEndpoint -Method POST -Headers $headers -Body $body
    
    # If user did not specify an output file, use the source report's name
    $outfile = $outfile ?? "$( `
      (Get-PowerBIReport -Id $sourceReportId -WorkspaceId $sourceWorkspaceId).Name `
      )_Clone.pbix"
    
    # Export the target report to a PBIX file
    Export-PowerBIReport -WorkspaceId $targetWorkspaceId -Id $response.id -OutFile $outFile
    
    # Assemble the Datasets API URI
    $datasetsEndpoint = "$pbiApiBaseUri/groups/$targetWorkspaceId/datasets"
    
    # Delete the target dataset and report from the target workspace
    Invoke-RestMethod "$datasetsEndpoint/$($response.datasetId)" -Method DELETE -Headers $headers
    
  }
  
}
