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

Function Get-PowerBIBareDatasetsFromWorkspaces {
	
	[CmdletBinding()]
	Param (
		[Parameter()][int]$ThrottleLimit = [Environment]::ProcessorCount,
		[Parameter()][switch]$Interactive
	)
	
	begin {
		# PowerShell Module dependencies
		#Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

		# Declare the servicePrincipal global variables
		# TODO: replace these with parameters to allow service principal authentication
		$global:servicePrincipalId = $null
    $global:servicePrincipalTenantId = $null
		$global:servicePrincipalSecret = $null
		# TODO: rewrite this line to use the PSCredential constructor instead of the New-Object cmdlet
		$global:credential = $servicePrincipalId ? (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $servicePrincipalId, ($servicePrincipalSecret | ConvertTo-SecureString -AsPlainText -Force)) : $null

		# Get names of Workspaces and Reports to ignore from IgnoreWorkspacesAndReports.json file
		# Most of these are template apps and/or auto-generated by Microsoft
		[PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "IgnoreList.json") | ConvertFrom-Json
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
			|	Where-Object { $hash.Add($_.DatasetId) }
	
	}

	end {
		Write-Verbose "Total number of Bare Datasets: $($hash.Count)"
		
		# Clear the PowerShell session's memory
		[gc]::Collect()
	}

}