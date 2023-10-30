#Requires -PSEdition Core
Function Get-PowerBIBareDatasetsFromWorkspaces {
  <#
  .SYNOPSIS
    Function: Get-PowerBIBareDatasetsFromWorkspaces
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)
  
  .DESCRIPTION
    Get all "bare" Power BI Datasets (Datasets without a corresponding report) from selected Workspaces in parallel
  
  .PARAMETER ThrottleLimit
    The maximum number of parallel processes to run.
    Defaults to 1.
  
  .PARAMETER Interactive
    If specified, displays a grid view of Workspaces and allows the user to select which ones to scan for bare Datasets.
  
  .INPUTS
    This function does not accept pipeline input.
  
  .OUTPUTS
    Selected.System.String (one or more objects with the following properties):
      - DatasetName
      - DatasetId
      - WebUrl
      - IsRefreshable
      - WorkspaceName
      - WorkspaceId
  
  .EXAMPLE
    Get-PowerBIBareDatasetsFromWorkspaces -Interactive -ThrottleLimit 4
  
  .LINK
    https://github.com/JamesDBartlett3/PowerBits
  
  .LINK
    https://techhub.social/@JamesDBartlett3
  
  .LINK
    https://datavolume.xyz
  
  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
  
    TODO
      - Separate verbose and debug outputs
      - HelpMessage on all parameters (https://youtu.be/UnjKVanzIOk)
      - 429 throttling (see Rui's repo and this article: https://powerbi.microsoft.com/en-us/blog/best-practices-to-prevent-getgroupsasadmin-api-timeout/)
      - Individual Datasets within a Workspace
      - Error handling and logging
      - Call Power BI REST API endpoints directly instead of MicrosoftPowerBIMgmt cmdlets
      - Service Principal authentication
      - [gc]::Collect() to free up memory
      - Testing
  
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to @santisq & @seeminglyscience on PowerShell Discord for their guidance on using 
        Hashset<T>.Add() to filter out duplicates in the output.
      - Thanks to @ruiromano on GitHub for his pbiscripts repo (https://github.com/RuiRomano/pbiscripts), 
        which inspired me, and taught me a lot about making Power BI REST API calls from PowerShell. 
        Much of the code in this repo is based on Rui's work.
  #>
  # PowerShell dependencies
  #Requires -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools
  [CmdletBinding()]
  Param (
    [Parameter()][int]$ThrottleLimit = [Environment]::ProcessorCount,
    [Parameter()][switch]$Interactive
  )
  begin {
    # Declare the servicePrincipal global variables
    # TODO: replace these with parameters to allow service principal authentication
    $global:servicePrincipalId = $null
    $global:servicePrincipalTenantId = $null
    $global:servicePrincipalSecret = $null
    # TODO: rewrite this line to use the PSCredential constructor instead of the New-Object cmdlet
    $global:credential = $servicePrincipalId ? (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $servicePrincipalId, ($servicePrincipalSecret | ConvertTo-SecureString -AsPlainText -Force)) : $null
    # Get names of Workspaces and Reports to ignore from IgnoreList.json file
    # Most of these are template apps and/or auto-generated by Microsoft
    [PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "../IgnoreList.json") | ConvertFrom-Json
    [array]$global:ignoreWorkspaces = $ignoreObjects.IgnoreWorkspaces
    [array]$global:ignoreReports = $ignoreObjects.IgnoreReports
  }
  process {
    try {
      $headers = Get-PowerBIAccessToken
    }
    catch {
      if ($servicePrincipalId) {
        $headers = Connect-PowerBIServiceAccount -ServicePrincipal -Tenant $servicePrincipalTenantId -Credential $credential
      } else {
        Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
        Start-Sleep -s 1
        Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
        $headers = Get-PowerBIAccessToken
      }
      if ($headers) {
        Write-Host '🔑 Power BI Access Token acquired. Proceeding...'
      } else {
        Write-Host '❌ Power BI Access Token not acquired. Exiting...'
        exit
      }
    }
    # Get the access token payload and convert it to JSON
    $token = $headers['Authorization'].Split(' ')[1]
    $tokenPayload = $token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    while ($tokenPayload.Length % 4) { $tokenPayload += '=' }
    $tokenPayload = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tokenPayload)) | ConvertFrom-Json
    # Get the user's UPN or Object ID, depending on whether the user is a service principal or not
    $pbiUserIdentifier = $servicePrincipalId ? $tokenPayload.oid : $tokenPayload.upn
    
    # If debugging, display the access token and user identifier
    Write-Debug "Headers: `n $($headers.Keys)`n $($headers.Values)`n"
    Write-Debug "User Identifier: `n $pbiUserIdentifier"
  
    # Get list of Workspaces
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All -ErrorAction SilentlyContinue | 
      Where-Object {
        $_.Type -eq 'Workspace' -and
        $_.State -eq 'Active' -and
        $_.Name -notIn $ignoreWorkspaces
      } | Select-Object Name, Id | Sort-Object -Property Name
  
    # If interactive, display a grid view of Workspaces and allow the user to select which ones to scan for bare Datasets
    $workspaces = $Interactive ? ($workspaces | Out-ConsoleGridView -Title 'Select Workspaces to Scan') : $workspaces
  
    # Declare $hash as a hashset to store unique Dataset IDs (prevents duplicates in the output)
    $hash = [System.Collections.Generic.Hashset[guid]]::New()
  
    # For each Workspace, find Datasets with no corresponding report and add them to the $bareDatasets array
    $workspaces | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
  
      # Declare local variables
      $workspaceName = $_.Name
      $workspaceId = $_.Id
    
      # Get Datasets from the Workspace
      $workspaceDatasets = Get-PowerBIDataset -Scope Organization -WorkspaceId $workspaceId -ErrorAction SilentlyContinue |
        Where-Object {
          $_.IsRefreshable -eq $true -and
          $_.Name -notIn $ignoreReports
        } | Select-Object Name, Id, WebUrl, IsRefreshable, @{
          Name = 'WorkspaceName'; Expression = { $workspaceName }
        }, @{
          Name = 'WorkspaceId'; Expression = { $workspaceId }
        } | Sort-Object -Property Name
    
        # Get reports from the Workspace
        $workspaceReports = Get-PowerBIReport -Scope Organization -WorkspaceId $workspaceId -ErrorAction SilentlyContinue |
          Where-Object {
            $_.Name -notIn $ignoreReports -and
            $_.WebUrl -notlike '*/rdlreports/*'
          } | Select-Object Name, Id, WebUrl, ReportType, DatasetId, @{
            Name = 'WorkspaceName'; Expression = { $workspaceName }
          }, @{
            Name = 'WorkspaceId'; Expression = { $workspaceId }
          } | Sort-Object -Property Name
    
          # For each Dataset, check for any corresponding reports with the same name
          $workspaceDatasets | ForEach-Object {
            $datasetProperties = '' | Select-Object DatasetName, DatasetId, WebUrl, IsRefreshable, WorkspaceName, WorkspaceId
            $datasetName, $datasetId, $datasetWebUrl, $datasetIsRefreshable, $datasetWorkspaceName, $datasetWorkspaceId = $null
            $datasetName = $_.Name
            $datasetId = $_.Id
            $datasetWebUrl = $_.WebUrl
            $datasetIsRefreshable = $_.IsRefreshable
            $datasetWorkspaceName = $_.WorkspaceName
            $datasetWorkspaceId = $_.WorkspaceId
      
            # If no corresponding report is found, output the Dataset's properties for processing downstream
            if (!($workspaceReports | Where-Object { $_.Name -eq $datasetName -and $_.WorkspaceId -eq $datasetWorkspaceId })) {
              $datasetProperties.DatasetName = $datasetName
              $datasetProperties.DatasetId = $datasetId
              $datasetProperties.WebUrl = $datasetWebUrl
              $datasetProperties.IsRefreshable = $datasetIsRefreshable
              $datasetProperties.WorkspaceName = $datasetWorkspaceName
              $datasetProperties.WorkspaceId = $datasetWorkspaceId
              return $datasetProperties
            }
      
          }
    
        } <# 
    Check each returned object for uniqueness by adding its DatasetId property to the hashset.
    If the DatasetId is not already in the hashset, $hash.Add() will return true, and the object 
    will be returned in the output. But if the DatasetId has already been added to the hashset, 
    $hash.Add() will return false, and the duplicate object will not be returned in the output.
    #> 
        | Where-Object { $hash.Add($_.DatasetId) }
  }
  end {
    Write-Verbose "Total number of Bare Datasets: $($hash.Count)"
  
    # Clear the PowerShell session's memory
    [gc]::Collect()
  }
}
Function Export-PowerBIBareDatasetsFromWorkspaces {
  <#
	.SYNOPSIS
		Function: Export-PowerBIBareDatasetsFromWorkspaces
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
	.INPUTS
		Selected.System.String (one or more objects with the following property names):
			- "DatasetId" or "Id" (required)
			- "WorkspaceId" (required)
			- "DatasetName" or "Name" (optional)
			- "WorkspaceName" (optional)
	.OUTPUTS
		This function does not output anything to the pipeline
	
	.EXAMPLE
		# Export a single Bare Dataset as a PBIX file by specifying the DatasetId, WorkspaceId, BlankPbix, and OutFile parameters
		Export-PowerBIBareDatasetsFromWorkspaces -DatasetId "00000000-0000-0000-0000-000000000000" -WorkspaceId "00000000-0000-0000-0000-000000000000" -BlankPbix "C:\blank.pbix" -OutFile "C:\new.pbix"
	
	.EXAMPLE 
		# Get a list of Bare Datasets from the Get-PowerBIBareDatasetsFromWorkspaces function
		$bareDatasets = Get-PowerBIBareDatasetsFromWorkspaces -Interactive
		# Then export them all as PBIX files
		$bareDatasets | Export-PowerBIBareDatasetsFromWorkspaces
	.LINK
		https://github.com/JamesDBartlett3/PowerBits
	.LINK
		https://techhub.social/@JamesDBartlett3
	.LINK
		https://datavolume.xyz
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

  #Requires -Modules MicrosoftPowerBIMgmt
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
    [Parameter(ValueFromPipelineByPropertyName)]
    [Alias('Name')][string]$DatasetName,
    [Parameter(ValueFromPipelineByPropertyName)]
    [string]$WorkspaceName,
    [Parameter()][string]$BlankPbix,
    [Parameter()][string]$OutFile
  )
  begin {
    [string]$blankPbixUri = 'https://github.com/JamesDBartlett3/PowerBits/raw/main/Misc/blank.pbix'
    [string]$tempFolder = Join-Path -Path $env:TEMP -ChildPath 'PowerBIBareDatasets'
    [string]$blankPbixTempFile = Join-Path -Path $env:TEMP -ChildPath 'blank.pbix'
    [string]$pbiApiBaseUri = 'https://api.powerbi.com/v1.0/myorg'
    [string]$urlRegex = '(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)'
    [array]$validPbixContents = @('Layout', 'Metadata')
    [bool]$blankPbixIsUrl = $BlankPbix -Match $urlRegex
    [bool]$localFileExists = Test-Path $BlankPbix
    [bool]$defaultFileIsValid = $false
    [bool]$remoteFileIsValid = $false
    [bool]$localFileIsValid = $false
    [int]$bareDatasetCount = 0
    [int]$errorCount = 0
    Invoke-Item -Path $tempFolder
  }
  process {
    Write-Debug "DatasetId: $DatasetId, WorkspaceId: $WorkspaceId, DatasetName: $DatasetName, WorkspaceName: $WorkspaceName"
	
    $headers = [System.Collections.Generic.Dictionary[[String],[String]]]::New()
	
    [string]$uniqueName = 'temp_' + [guid]::NewGuid().ToString().Replace('-', '')
	
    Function FileIsBlankPbix($file) {
      $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
      $fileIsPbix = @($validPbixContents | Where-Object { $zip.Entries.Name -Contains $_ }).Count -gt 0
      $fileIsBlank = (Get-Item $file).length / 1KB -lt 20
      $zip.Dispose()
      if ($fileIsPbix -and $fileIsBlank) {
        Write-Verbose "$file is a valid blank PBIX file."
        return $true
      } else {
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
      Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
      Start-Sleep -s 1
      Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
      $headers = Get-PowerBIAccessToken
      if ($headers) {
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
    # Rebind the published Report to the Bare Dataset
    Write-Verbose "Rebinding published Report $publishedReportId to Dataset $DatasetId..."
    Invoke-RestMethod -Uri $updateReportContentEndpoint -Method POST -Headers $headers -Body $body | Out-Null
    # If the Workspace folder doesn't exist, create it
    if (!(Test-Path (Join-Path -Path $tempFolder -ChildPath $WorkspaceName))) {
      New-Item -Path (Join-Path -Path $tempFolder -ChildPath $WorkspaceName) -ItemType Directory | Out-Null
    }
    # If user did not specify an output file name, use the Dataset's name and save it in the default temp folder
    $OutFile = if (!!$OutFile) { $OutFile } else {
      Join-Path -Path $tempFolder -ChildPath (Join-Path -Path $WorkspaceName -ChildPath "$($DatasetName).pbix")
    }
    # Export the re-bound Report and Dataset (a.k.a. "Thick Report") PBIX file to a temp file first, then rename it to the correct name (workaround for Datasets with special characters in their names)
    Write-Verbose "Exporting re-bound blank Report and Dataset (a.k.a. 'Thick Report') $publishedReportId to temporary file $($uniqueName).pbix..."
    $tempFileName = Join-Path -Path $tempFolder -ChildPath "$uniqueName.pbix"
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
      Write-Error "Error exporting Bare Dataset `"$DatasetName`" from `"$WorkspaceName`"`: $errorCode"
    } else {
      $bareDatasetCount++
      Write-Verbose "Exported Bare Dataset `"$DatasetName`" from `"$WorkspaceName`" to $tempFileName"
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
	
    # Remove any empty directories
    Get-ChildItem $tempFolder -Recurse -Attributes Directory | Where-Object { $_.GetFileSystemInfos().Count -eq 0 } | Remove-Item
	
    # Clear the PowerShell session's memory
    [gc]::Collect()
  }
}
