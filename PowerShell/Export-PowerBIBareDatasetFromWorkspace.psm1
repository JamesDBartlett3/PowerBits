<#
	.SYNOPSIS
		Function: Export-PowerBIBareDatasetFromWorkspace
		Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

	.DESCRIPTION
		Exports Bare Dataset (Dataset with no corresponding Report) from Power BI as PBIX file

	.PARAMETER DatasetId
		The ID of the Dataset to export

	.PARAMETER WorkspaceId
		The ID of the Workspace containing the Dataset to export

	.PARAMETER DatasetName
		The name of the Dataset to export

	.PARAMETER WorkspaceName
		The name of the Workspace containing the Dataset to export

	.PARAMETER BlankPbix
		Path (local or URL) to a blank PBIX file to upload and rebind to the Dataset to be exported

	.PARAMETER OutFile
		Local path to save the Dataset PBIX file to

	.EXAMPLE
		Export-PowerBIBareDatasetFromWorkspace -DatasetId "00000000-0000-0000-0000-000000000000" -WorkspaceId "00000000-0000-0000-0000-000000000000" -BlankPbix "C:\blank.pbix" -OutFile "C:\new.pbix"

	.NOTES
		This function does NOT require Azure AD app registration, 
		service principal creation, or any other special setup.
		The only requirements are:
			- The user must be able to run PowerShell (and install the
				MicrosoftPowerBIMgmt module, if it's not already installed).
			- The user must be allowed to download report PBIX files
				(see: "Download reports" setting in the Power BI Admin Portal).
			- The user must have "Contributor" or higher permissions on the 
				Workspace(s) where the Bare Dataset(s) to be exported are published.

		TODO
			- Error handling and logging
			- Parallelism
			- ParameterSetName on mutually-exclusive parameters (https://youtu.be/OO2yu5RgOVo)
			- HelpMessage on all parameters (https://youtu.be/UnjKVanzIOk)
			- [ValidateScript({Test-Path $_})][string]$path on all file paths
			- 429 throttling (see Rui's repo and this article: https://powerbi.microsoft.com/en-us/blog/best-practices-to-prevent-getgroupsasadmin-api-timeout/)
			- Call Power BI REST API endpoints directly instead of MicrosoftPowerBIMgmt cmdlets
			- Service Principal authentication
			- [gc]::Collect() to free up memory
			- Testing

		ACKNOWLEDGEMENTS
			- Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
			- Thanks to @santisq & @seeminglyscience on PowerShell Discord for their guidance on using 
				a process block to enable streaming inputs from the pipeline.
#>

Function Export-PowerBIBareDatasetFromWorkspace {

	[CmdletBinding()]
	Param(
		[Parameter(
			Mandatory,
			ValueFromPipelineByPropertyName
		)][Alias('Id')][guid]$DatasetId,
		[Parameter(
			Mandatory,
			ValueFromPipelineByPropertyName
		)][guid]$WorkspaceId,
		[Parameter(
			, ValueFromPipelineByPropertyName
		)][Alias('Name')][string]$DatasetName,
		[Parameter(
			, ValueFromPipelineByPropertyName
		)][string]$WorkspaceName,
		[Parameter()][string]$BlankPbix,
		[Parameter()][string]$OutFile
	)

	begin {
		#Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt
		[int]$bareDatasetCount = 0
		[int]$errorCount = 0
	}

	process {

		Write-Debug "DatasetId: $DatasetId, WorkspaceId: $WorkspaceId, DatasetName: $DatasetName, WorkspaceName: $WorkspaceName"
		
		$headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
		
		[string]$tempFolder = Join-Path -Path $env:TEMP -ChildPath 'PowerBIBareDatasets'
		[string]$blankPbixTempFile = Join-Path -Path $env:TEMP -ChildPath 'blank.pbix'
		[string]$urlRegex = '(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)'
		[string]$uniqueName = 'temp_' + [guid]::NewGuid().ToString().Replace('-', '')
		[string]$pbiApiBaseUri = 'https://api.powerbi.com/v1.0/myorg'
		[array]$validPbixContents = @('Layout', 'Metadata')
		[bool]$blankPbixIsUrl = $BlankPbix -Match $urlRegex
		[bool]$localFileExists = Test-Path $BlankPbix
		[bool]$remoteFileIsValid = $false
		[bool]$localFileIsValid = $false
		[bool]$defaultFileIsValid = $false
		
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

		# If the temp folder doesn't exist and user has not specified OutFile location, create the temp folder
		if (!(Test-Path -LiteralPath $tempFolder) -and !$OutFile) {
			New-Item -Path $tempFolder -ItemType Directory | Out-Null
		}
		
		# If user specified a URL to a file, download and validate it as a blank PBIX file
		if ($blankPbixIsUrl) {
			Write-Verbose "Downloading file: $BlankPbix..."
			Invoke-WebRequest -Uri $BlankPbix -OutFile $blankPbixTempFile
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
			$BlankPbixUri = 'https://github.com/JamesDBartlett3/PowerBits/raw/main/Misc/blank.pbix'
			Invoke-WebRequest -Uri $BlankPbixUri -OutFile $blankPbixTempFile
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
			Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
			Start-Sleep -s 1
			Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
			$headers = Get-PowerBIAccessToken
			if($headers) {
				Write-Host '🔑 Power BI Access Token acquired. Proceeding...'
			} else {
				Write-Host '❌ Power BI Access Token not acquired. Exiting...'
				Exit
			}
		}

		# If user did not specify a Dataset name, get it from the API
		$DatasetName = $DatasetName ?? (Get-PowerBIDataset -Id $DatasetId -WorkspaceId $WorkspaceId).Name
		
		# If user did not specify a Workspace name, get it from the API
		$WorkspaceName = $WorkspaceName ?? (Get-PowerBIWorkspace -Id $WorkspaceId).Name

		# Publish the blank PBIX file to the target workspace
		Write-Verbose "Publishing $BlankPbix to `"$WorkspaceName`" Workspace with temporary name $uniqueName"
		$publishResponse = New-PowerBIReport -Path $BlankPbix -WorkspaceId $WorkspaceId -Name $uniqueName -ConflictAction CreateOrOverwrite
		Write-Debug "Response: $publishResponse"
		$publishedReportId = $publishResponse.Id
		$publishedDatasetId = (Get-PowerBIDataset -WorkspaceId $WorkspaceId | Where-Object { $_.Name -eq $uniqueName }).Id
		Write-Debug "Published Report ID: $publishedReportId; Published Dataset ID: $publishedDatasetId"

		# Assemble the Workspace base URI
		$workspaceBaseUri = "$pbiApiBaseUri/groups/$WorkspaceId"
		# Assemble the Datasets API URI
		$datasetsEndpoint = "$workspaceBaseUri/datasets"
		# Assemble the Reports API URI
		$reportsEndpoint = "$workspaceBaseUri/reports"
		# Assemble the Rebind API URI
		$updateReportContentEndpoint = "$reportsEndpoint/$publishedReportId/Rebind"
		# Assemble the Rebind API request body
		$body = "{`"datasetId`": `"$DatasetId`"}"
		# Assemble the Export API URI
		$exportEndpoint = "$reportsEndpoint/$publishedReportId/Export"

		# Add the Content-Type header to the request
		$headers.Add('Content-Type', 'application/json')

		# Rebind the published Report to the bare Dataset
		Write-Verbose "Rebinding published Report $publishedReportId to Dataset $DatasetId..."
		Invoke-RestMethod -Uri $updateReportContentEndpoint -Method POST -Headers $headers -Body $body | Out-Null

		# If the Workspace folder doesn't exist, create it
		if (!(Test-Path (Join-Path -Path $tempFolder -ChildPath $WorkspaceName))) {
			New-Item -Path (Join-Path -Path $tempFolder -ChildPath $WorkspaceName) -ItemType Directory | Out-Null
		}

		# If user did not specify an output file name, use the Dataset's name and save it in the default temp folder
		$OutFile = if (!!$OutFile) { $OutFile } else {
			Join-Path -Path $tempFolder -ChildPath (Join-Path -Path $WorkspaceName -ChildPath "$($DatasetName).pbix")
			Invoke-Item -Path $tempFolder
		}

		# Export the re-bound Report and Dataset (a.k.a. "Thick Report") PBIX file to a temp file first, then rename it to the correct name (workaround for Datasets with special characters in their names)
		Write-Verbose "Exporting re-bound blank Report and Dataset (a.k.a. 'Thick Report') $publishedReportId to temporary file $($uniqueName).pbix..."
		$tempFileName = Join-Path -Path $tempFolder -ChildPath "$uniqueName.pbix"
		$message = Invoke-RestMethod -Uri "$exportEndpoint" `
			-Method GET -Headers $headers `
			-ContentType "application/octet-stream" `
			-Body '{"preferClientRouting":true}' `
			-ErrorVariable message `
			-ErrorAction SilentlyContinue `
			-OutFile $tempFileName 2>&1 | Out-String
		$message = switch ($true) {
			{ $message -like "*BadRequest*" } { "Incremental Refresh." }
			{ $message -like "*NotFound*" -or $message -like "*Forbidden*" -or $message -like "*Disabled*" } { "Downloads Disabled." }
			{ $message -like "*TooManyRequests*" } { "Reached Power BI API Rate Limit; Try Again Later." }
			{ $message -like "*Unauthorized*" } { "Unauthorized." }
			default { $null }
		}
		if($message) {
			$errorCount++
			Write-Error "Error exporting bare Dataset `"$DatasetName`" from `"$WorkspaceName`"`: $message"
		} else {
			$bareDatasetCount++
			Write-Verbose "Exported bare Dataset `"$DatasetName`" from `"$WorkspaceName`" to $tempFileName"
			Write-Verbose "Moving and renaming temp file $($uniqueName).pbix to $OutFile..."
			Move-Item -Path $tempFileName -Destination $OutFile -Force
		}
		$OutFile = $null

		# Delete the blank Report and its original Dataset from the Workspace
		Write-Verbose "Deleting temporary blank Dataset $publishedDatasetId and Report $publishedReportId from Workspace $WorkspaceId..."
		Invoke-RestMethod "$datasetsEndpoint/$publishedDatasetId" -Method DELETE -Headers $headers | Out-Null
		Invoke-RestMethod "$reportsEndpoint/$publishedReportId" -Method DELETE -Headers $headers | Out-Null
		
	}

	end {
		Write-Verbose "Bare Datasets successfully exported: $bareDatasetCount.$(if($errorCount -gt 0){" Errors encountered: $errorCount"})"
		[gc]::Collect()
	}

}