<#
  .SYNOPSIS
    Function: Export-PowerBIBareDatasetsFromWorkspaces
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

  .DESCRIPTION
    - Exports Bare Datasets (Datasets with no corresponding Reports) from Power BI as PBIX files

  .PARAMETER DatasetId
    The ID of the Dataset to export

  .PARAMETER WorkspaceId
    The ID of the workspace containing the Dataset to export

  .PARAMETER BlankPbix
    Path (local or URL) to a blank PBIX file to upload and copy the source report's contents into

  .PARAMETER OutFile
    Local path to save the dataset PBIX file to

  .EXAMPLE
    Export-PowerBIBareDatasetsFromWorkspaces -DatasetId "00000000-0000-0000-0000-000000000000" -WorkspaceId "00000000-0000-0000-0000-000000000000" -BlankPbix "C:\blank.pbix" -OutFile "C:\new.pbix"

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
    
    TODO
			- Add support for multiple datasets and workspaces
			- Refactor to use the Power BI REST API directly instead of the MicrosoftPowerBIMgmt cmdlets
      - Testing
  
    ACKNOWLEDGEMENTS
			- Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
#>

Function Export-PowerBIBareDatasetsFromWorkspaces {
  
  #Requires -PSEdition Core
  #Requires -Modules MicrosoftPowerBIMgmt
  
  [CmdletBinding()]
  Param(
    [parameter(Mandatory = $true)][string]$DatasetId,
    [parameter(Mandatory = $true)][string]$WorkspaceId,
    [Parameter(Mandatory = $false)][string]$BlankPbix,
    [Parameter(Mandatory = $false)][string]$OutFile
  )
  
  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  
	[string]$tempFolder = Join-Path -Path $env:TEMP -ChildPath "PowerBIBareDatasets"
  [string]$blankPbixTempFile = Join-Path -Path $env:TEMP -ChildPath "blank.pbix"
  [array]$validPbixContents = @("Layout", "Metadata")
	[string]$urlRegex = "(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)"
	[string]$uniqueName = "temp_" + [guid]::NewGuid().ToString().Replace("-","")
  [bool]$blankPbixIsUrl = $BlankPbix -Match $urlRegex
  [bool]$localFileExists = Test-Path $BlankPbix
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
      Write-Error "$file is NOT a valid pbix file and/or NOT blank."
      return $false
    }
  }

	# If the temp folder doesn't exist and user has not specified OutFile location, create the temp folder
	if (!(Test-Path -LiteralPath $tempFolder) -and !$OutFile) {
		New-Item -Path $tempFolder -ItemType Directory | Out-Null
	}
  
	# If user specified a URL to a file, download and validate it as a blank PBIX file
	if ($blankPbixIsUrl){
		Write-Debug "Downloading file: $BlankPbix..."
		Invoke-WebRequest -Uri $BlankPbix -OutFile $blankPbixTempFile
		Write-Debug "Validating downloaded file..."
		$remoteFileIsValid = FileIsBlankPbix($blankPbixTempFile)
	}
	
	# If user specified a local path to a file, validate it as a blank PBIX file
	elseif ($localFileExists) {
		Write-Debug "Validating user-supplied file: $BlankPbix..."
		$localFileIsValid = FileIsBlankPbix($BlankPbix)
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
		$BlankPbixUri = "https://github.com/JamesDBartlett3/PowerBits/raw/main/Misc/blank.pbix"
		Invoke-WebRequest -Uri $BlankPbixUri -OutFile $blankPbixTempFile
		$defaultFileIsValid = FileIsBlankPbix($blankPbixTempFile)
	}
	
	# If we downloaded a valid blank PBIX file, use it.
	if ($remoteFileIsValid -or $defaultFileIsValid) {
		$BlankPbix = $blankPbixTempFile
	}
	
	# If a valid blank PBIX file could not be obtained by any of the above methods, throw an error.
	if (!$localFileIsValid -and !$remoteFileIsValid -and !$defaultFileIsValid) {
		Write-Error "No blank PBIX file found. Please specify a valid blank PBIX file using the -BlankPbix parameter."
		return
	}
  
  try {
    $headers = Get-PowerBIAccessToken
  } catch {
    Write-Host "🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog..."
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
  } finally {
    Write-Host "🔑 Power BI Access Token acquired."
    $pbiApiBaseUri = "https://api.powerbi.com/v1.0/myorg"
    
		# Publish the blank PBIX file to the target workspace
		Write-Debug "Publishing $BlankPbix to workspace with temporary name $uniqueName"
		$publishResponse = New-PowerBIReport -Path $BlankPbix -WorkspaceId $WorkspaceId -Name $uniqueName -ConflictAction CreateOrOverwrite
		Write-Debug "Response: $publishResponse"
		$publishedReportId = $publishResponse.Id
		$publishedDatasetId = (Get-PowerBIDataset -WorkspaceId $WorkspaceId | Where-Object {$_.Name -eq $uniqueName}).Id
		Write-Debug "Published report ID: $publishedReportId; Published dataset ID: $publishedDatasetId"

		# Assemble the Datasets API URI
		$datasetsEndpoint = "$pbiApiBaseUri/groups/$WorkspaceId/datasets"
		# Assemble the Reports API URI
		$reportsEndpoint = "$pbiApiBaseUri/groups/$WorkspaceId/reports"
    # Assemble the Rebind API URI and request body
    $updateReportContentEndpoint = "$pbiApiBaseUri/groups/$WorkspaceId/reports/$publishedReportId/Rebind"
    $body = @"
      {
        "datasetId": "$DatasetId"
      }
"@
    $headers.Add("Content-Type", "application/json")

		# Rebind the published report to the bare dataset
		Write-Debug "Rebinding published report $publishedReportId to dataset $DatasetId..."
    Invoke-RestMethod -Uri $updateReportContentEndpoint -Method POST -Headers $headers -Body $body
    
    # If user did not specify an output file name, use the dataset's name and save it in the default temp folder
    $OutFile = if(!!$OutFile) {$OutFile} else {
			Join-Path -Path $tempFolder -ChildPath "$((Get-PowerBIDataset -Id $DatasetId -WorkspaceId $WorkspaceId).Name).pbix"
			Invoke-Item -Path $tempFolder
		}
    
    # Export the re-bound report and dataset (a.k.a. "thick report") to a PBIX file
		Write-Debug "Exporting re-bound blank report and dataset (a.k.a. 'thick report') $publishedReportId to $OutFile..."
    Export-PowerBIReport -WorkspaceId $WorkspaceId -Id $publishedReportId -OutFile $OutFile
  
    # Delete the published blank dataset and report from the workspace
		Write-Debug "Deleting temporary blank dataset $publishedDatasetId and report $publishedReportId from workspace $WorkspaceId..."
    Invoke-RestMethod "$datasetsEndpoint/$publishedDatasetId" -Method DELETE -Headers $headers
		Invoke-RestMethod "$reportsEndpoint/$publishedReportId" -Method DELETE -Headers $headers
    
  }
  
}
