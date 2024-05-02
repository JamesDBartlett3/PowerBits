<#
  .SYNOPSIS
    Exports one or more "Thin" Models (Power BI Semantic Models with no corresponding Report) from Power BI as PBIX files.
  
  .DESCRIPTION
    This script will export one or more "Thin" Models (Power BI Semantic Models with no corresponding Report) from Power BI as PBIX files.
    It can be used to export a single Thin Model by specifying the DatasetId, WorkspaceId, BlankPbix, and OutputFile parameters.
    It can also be used to export multiple Thin Models by piping the output of the Get-PowerBIThinModelsFromWorkspaces.ps1 script to it.
  
  .PARAMETER DatasetId
    The ID of the Model to export
  
  .PARAMETER WorkspaceId
    The ID of the Workspace containing the Model to export
  
  .PARAMETER DatasetName
    The name of the Model to export
  
  .PARAMETER WorkspaceName
    The name of the Workspace containing the Model to export
  
  .PARAMETER OutputFile
    Local filename for one exported Model PBIX file
  
  .PARAMETER OutputFolder
    Local folder to store multiple workspace subfolders and exported Model PBIX files
  
  .PARAMETER BlankPbix
    Path (local or URL) to a blank PBIX file to upload and rebind to the Model to be exported
  
  .INPUTS
    Selected.System.String (one or more objects with the following property names):
      - "DatasetId" or "Id" (required)
      - "WorkspaceId" (required)
      - "DatasetName" or "Name" (optional)
      - "WorkspaceName" (optional)
  
  .OUTPUTS
    This script does not output anything to the pipeline
  
  .EXAMPLE
    # Export a single Thin Model as a PBIX file by specifying the DatasetId, WorkspaceId, BlankPbix, and OutputFile parameters
    .\Export-PowerBIThinModelsFromWorkspaces.ps1 -DatasetId "00000000-0000-0000-0000-000000000000" -WorkspaceId "00000000-0000-0000-0000-000000000000" -BlankPbix "C:\blank.pbix" -OutputFile "C:\new.pbix"
	
  .EXAMPLE 
    # Use the Get-PowerBIThinModelsFromWorkspaces.ps1 script with the Interactive switch to pick one or more Thin Models,
    # and pipe the output into Export-PowerBIThinModelsFromWorkspaces.ps1 to export the selected Thin Models as PBIX files
    .\Get-PowerBIThinModelsFromWorkspaces.ps1 -Interactive | .\Export-PowerBIThinModelsFromWorkspaces.ps1
  
  .EXAMPLE
    # Pipe the output of the Get-PowerBIThinModelsFromWorkspaces.ps1 script to this script, export all detected thin models to which the user has access,
    # and use the OutputFolder parameter to specify that they should be stored in a folder named "PowerBIThinModels" in the user's Downloads directory
    .\Get-PowerBIThinModelsFromWorkspaces.ps1 | .\Export-PowerBIThinModelsFromWorkspaces.ps1 -OutputFolder "$env:USERPROFILE\Downloads\PowerBIThinModels"
	
  .NOTES
    This script does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must be allowed to download report PBIX files
        (see: "Download reports" setting in the Power BI Admin Portal).
      - The user must have "Contributor" or higher permissions on the 
        Workspace(s) where the Thin Model(s) to be exported are published.
    
TODO
      - Error handling and logging
      - Parallelism
      - ParameterSetName on mutually-exclusive parameters (https://youtu.be/OO2yu5RgOVo)
      - HelpMessage on all parameters (https://youtu.be/UnjKVanzIOk)
      - [ValidateScript({Test-Path $_})][string]$path on all file paths
      - 429 throttling (see Rui's repo and this article: https://powerbi.microsoft.com/en-us/blog/best-practices-to-prevent-getgroupsasadmin-api-timeout/)
      - Call Power BI REST API endpoints directly instead of MicrosoftPowerBIMgmt cmdlets
      - Service Principal authentication
      - Testing
    
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
      - Thanks to @santisq & @seeminglyscience on PowerShell Discord for their guidance on using a process block to enable streaming inputs from the pipeline.
		
    TODO: Add ExtractWithPbiTools switch parameter
  
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits/blob/main/PowerShell/Scripts/Export-PowerBIThinModelsFromWorkspaces.ps1)
  
  .LINK
    [The author's blog](https://datavolume.xyz)
    
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
#>

#Requires -PSEdition Core
#Requires -Modules MicrosoftPowerBIMgmt

[CmdletBinding()]
Param(
  [Parameter(Mandatory, ValueFromPipelineByPropertyName)][Alias('Id')][guid]$DatasetId,
  [Parameter(Mandatory, ValueFromPipelineByPropertyName)][guid]$WorkspaceId,
  [Parameter(ValueFromPipelineByPropertyName)][Alias('Name')][string]$ModelName,
  [Parameter(ValueFromPipelineByPropertyName)][string]$WorkspaceName,
  [Parameter()][string]$OutputFolder,
  [Parameter()][string]$OutputFile,
  [Parameter()][string]$BlankPbix
)

begin {
  [string]$blankPbixUri = 'https://github.com/JamesDBartlett3/PowerBits/raw/main/Misc/blank.pbix'
  [string]$tempFolder = Join-Path -Path $env:TEMP -ChildPath 'PowerBIThinModels'
  [string]$outputDirectory = if(!($OutputFolder)) {$tempFolder} else {$OutputFolder}
  [string]$blankPbixTempFile = Join-Path -Path $env:TEMP -ChildPath 'blank.pbix'
  [string]$pbiApiBaseUri = 'https://api.powerbi.com/v1.0/myorg'
  [string]$urlRegex = '(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)'
  [array]$validPbixContents = @('Layout', 'Metadata')
  [bool]$blankPbixIsUrl = $BlankPbix -Match $urlRegex
  [bool]$localFileExists = Test-Path $BlankPbix
  [bool]$defaultFileIsValid = $false
  [bool]$remoteFileIsValid = $false
  [bool]$localFileIsValid = $false
  [int]$thinModelCount = 0
  [int]$errorCount = 0
	
  if(!(Test-Path -LiteralPath $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory | Out-Null
  }
	
  Invoke-Item -Path $outputDirectory
  Function FileIsBlankPbix($file) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
    $fileIsPbix = @($validPbixContents | Where-Object { $zip.Entries.Name -Contains $_ }).Count -gt 0
    $fileIsBlank = (Get-Item $file).length / 1KB -lt 20
    $zip.Dispose()
    if ($fileIsPbix -and $fileIsBlank) {
      Write-Verbose "$file is a valid blank PBIX file."
      return $true
    }
    else {
      Write-Error "$file is NOT a valid PBIX file and/or NOT blank."
      return $false
    }
  }
}

process {
  
  Write-Debug "DatasetId: $DatasetId, WorkspaceId: $WorkspaceId, DatasetName: $ModelName, WorkspaceName: $WorkspaceName"
  
  $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()
  
  [string]$uniqueName = 'temp_' + [guid]::NewGuid().ToString().Replace('-', '')
  
  # If the output folder doesn't exist and user has not specified OutputFile location, create the output folder
  if (!(Test-Path -LiteralPath $outputDirectory) -and !$OutputFile) {
    New-Item -Path $outputDirectory -ItemType Directory | Out-Null
  }
  
  # If user specified a URL to a file, download and validate it as a blank PBIX file
  if ($blankPbixIsUrl) {
    Write-Verbose "Downloading file: $BlankPbix..."
    Invoke-WebRequest -Uri $BlankPbix -OutputFile $blankPbixTempFile
    Write-Verbose 'Validating downloaded file...'
    $remoteFileIsValid = FileIsBlankPbix($blankPbixTempFile)
  }
  
  # If user specified a local path to a file, validate it as a blank PBIX file
  elseif ($localFileExists) {
    Write-Verbose "Validating user-supplied file: $BlankPbix..."
    $localFileIsValid = FileIsBlankPbix($BlankPbix)
  }
  
  # If user didn't specify a blank PBIX file, check for a valid blank PBIX in the temp location
  elseif (Test-Path $blankPbixTempFile) {
    Write-Verbose "Validating pbix file found in temp location: $blankPbixTempFile..."
    $defaultFileIsValid = FileIsBlankPbix($blankPbixTempFile)
  }
  
  # If user did not specify a blank PBIX file, and a valid blank PBIX is not in the temp location,
  # download one from GitHub, and then check if it's valid and blank
  else {
    Write-Verbose "Downloading a blank pbix file from GitHub to $blankPbixTempFile..."
    Invoke-WebRequest -Uri $blankPbixUri -OutFile $blankPbixTempFile
    $defaultFileIsValid = FileIsBlankPbix($blankPbixTempFile)
  }
  
  # If we downloaded a valid blank PBIX file, use it.
  if ($remoteFileIsValid -or $defaultFileIsValid) {
    $BlankPbix = $blankPbixTempFile
  }
  
  # If a valid blank PBIX file could not be obtained by any of the above methods, throw an error.
  if (!$localFileIsValid -and !$remoteFileIsValid -and !$defaultFileIsValid) {
    Write-Error 'No valid blank PBIX file found. Please specify a valid blank PBIX file using the -BlankPbix parameter.'
    return
  }
  
  try {
    $headers = Get-PowerBIAccessToken
  }
  
  catch {
    Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID (f.k.a. Azure Active Directory) authentication dialog...'
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
    if ($headers) {
      Write-Host '🔑 Power BI Access Token acquired. Proceeding...'
    }
    else {
      Write-Host '❌ Power BI Access Token not acquired. Exiting...'
      Exit
    }
  }
  
  # Add Content-Type: application/json to the headers
  $headers.Add('Content-Type', 'application/json')
  
  # If user did not specify a Model name, get it from the API
  $ModelName = $ModelName ?? (Get-PowerBIDataset -Id $DatasetId -WorkspaceId $WorkspaceId).Name
  
  # If user did not specify a Workspace name, get it from the API
  $WorkspaceName = $WorkspaceName ?? (Get-PowerBIWorkspace -Id $WorkspaceId).Name
  
  # Publish the blank PBIX file to the target workspace
  Write-Verbose "Publishing $BlankPbix to `"$WorkspaceName`" Workspace with temporary name $uniqueName"
  $publishResponse = New-PowerBIReport -Path $BlankPbix -WorkspaceId $WorkspaceId -Name $uniqueName -ConflictAction CreateOrOverwrite
  Write-Debug "Response: $publishResponse"
  $publishedReportId = $publishResponse.Id
  $publishedDatasetId = (Get-PowerBIDataset -WorkspaceId $WorkspaceId | Where-Object { $_.Name -eq $uniqueName }).Id
  Write-Debug "Published Report ID: $publishedReportId; Published Model ID: $publishedDatasetId"
  
  # Assemble the Workspace base URI
  [string]$workspaceBaseUri = "$pbiApiBaseUri/groups/$WorkspaceId"
  # Assemble the Datasets API URI
  [string]$datasetsEndpoint = "$workspaceBaseUri/datasets"
  # Assemble the Reports API URI
  [string]$reportsEndpoint = "$workspaceBaseUri/reports"
  # Assemble the Rebind API URI
  [string]$updateReportContentEndpoint = "$reportsEndpoint/$publishedReportId/Rebind"
  # Assemble the Export API URI
  [string]$exportEndpoint = "$reportsEndpoint/$publishedReportId/Export"
  # Assemble the Rebind API request body
  [string]$body = "{`"datasetId`": `"$DatasetId`"}"
  
  # Rebind the published Report to the Thin Model
  Write-Verbose "Rebinding published Report $publishedReportId to Model $DatasetId..."
  Invoke-RestMethod -Uri $updateReportContentEndpoint -Method POST -Headers $headers -Body $body | Out-Null
  
  # If the Workspace folder doesn't exist, create it
  if (!(Test-Path (Join-Path -Path $outputDirectory -ChildPath $WorkspaceName))) {
    New-Item -Path (Join-Path -Path $outputDirectory -ChildPath $WorkspaceName) -ItemType Directory | Out-Null
  }
  
  # If user did not specify an output file name, use the Model's name and save it in the output folder
  $OutputFile = if (!$OutputFile) {
    Join-Path -Path $outputDirectory -ChildPath (Join-Path -Path $WorkspaceName -ChildPath "$($ModelName).pbix") 
  } else {
    $OutputFile
  }
  
  # Export the re-bound Report and Model (a.k.a. "Thick Report") PBIX file to a temp file first, then 
  # rename it to the correct name (workaround for Models with special characters in their names)
  Write-Verbose "Exporting re-bound blank Report and Model (a.k.a. 'Thick Report') $publishedReportId to temporary file $($uniqueName).pbix..."
  [string]$tempFileName = Join-Path -Path $outputDirectory -ChildPath "$uniqueName.pbix"
  Invoke-RestMethod -Uri "$exportEndpoint" `
    -Method GET -Headers $headers `
    -ContentType 'application/octet-stream' `
    -Body '{"preferClientRouting":true}' `
    -ErrorVariable message `
    -ErrorAction SilentlyContinue `
    -OutFile $tempFileName 2>&1 | Out-Null
  if ($message) {
    $errorCount++
    $errorCode = ($message.ErrorRecord.ErrorDetails.Message | ConvertFrom-Json).error.code
    Write-Error "Error exporting Thin Model `"$ModelName`" from `"$WorkspaceName`"`: $errorCode"
  }
  else {
    $thinModelCount++
    Write-Verbose "Exported Thin Model `"$ModelName`" from `"$WorkspaceName`" to $tempFileName"
    Write-Verbose "Moving and renaming temp file $($uniqueName).pbix to $OutputFile..."
    Move-Item -Path $tempFileName -Destination $OutputFile -Force
  }
  $OutputFile = $null
  
  # Delete the blank Report and its original Model from the Workspace
  Write-Verbose "Deleting temporary blank Model $publishedDatasetId and Report $publishedReportId from Workspace $WorkspaceId..."
  Invoke-RestMethod "$datasetsEndpoint/$publishedDatasetId" -Method DELETE -Headers $headers | Out-Null
  Invoke-RestMethod "$reportsEndpoint/$publishedReportId" -Method DELETE -Headers $headers | Out-Null
  
}

end {
  Write-Verbose "Thin Models successfully exported: $thinModelCount -- $(if($errorCount -gt 0){" Errors encountered: $errorCount"})"
  
  # Remove any empty directories
  Get-ChildItem $outputDirectory -Recurse -Attributes Directory | Where-Object { $_.GetFileSystemInfos().Count -eq 0 } | Remove-Item
  
  # Clear the PowerShell session's memory
  [gc]::Collect()
}