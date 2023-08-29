<#
	
.SYNOPSIS
	Function: Get-PowerBIBareDatasetsFromWorkspaces
	Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

.DESCRIPTION
	Get all "bare" Power BI datasets (datasets without a corresponding report) from selected workspaces in parallel

.EXAMPLE
	Get-PowerBIBareDatasetsFromWorkspaces -ThrottleLimit 10

.PARAMETER ThrottleLimit
	The maximum number of parallel processes to run.
	Defaults to 1.

.PARAMETER Interactive
	If specified, displays a grid view of workspaces and allows the user to select which ones to scan for bare datasets.

.NOTES
	This function does NOT require Azure AD app registration, 
	service principal creation, or any other special setup.
	The only requirements are:
		- The user must be able to run PowerShell (and install the
			MicrosoftPowerBIMgmt module, if it's not already installed).

	TODO
		- 429 throttling
		- Individual datasets within a workspace
		- Pipeline streaming
		- Error handling and logging
		- Call Power BI REST API endpoints directly instead of MicrosoftPowerBIMgmt cmdlets
		- Testing

	ACKNOWLEDGEMENTS
		- Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
#>

Function Get-PowerBIBareDatasetsFromWorkspaces {
	
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $false)][int]$ThrottleLimit = 1,
		[Parameter(Mandatory = $false)][switch]$Interactive
	)
	
	#Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools
	
	try {
		$headers = Get-PowerBIAccessToken
	} 
	catch {
		Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
		Start-Sleep -s 1
		Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
		$headers = Get-PowerBIAccessToken
		Write-Host '🔑 Power BI Access Token acquired. Proceeding...'
	} 
	finally {
		
		# If debugging, display the access token
		Write-Debug "Headers: `n $($headers.Keys)`n $($headers.Values)"
		
		# Define names of workspaces and reports to ignore
		# Most of these are template apps and/or auto-generated by Microsoft
		[array]$ignoreWorkspaces = @(
			'Apps Catalog on Microsoft AppSource'
			, 'Azure DevOps Dashboard'
			, 'COVID-19 Global Report'
			, 'COVID-19 US Tracking Report'
			, 'Custom Visuals Exploration Tool'
			, 'Dataflow Snapshots'
			, 'Gen2 Utilization Metrics'
			, 'Microsoft 365 Usage Analytics'
			, 'Microsoft Fabric Capacity Metrics'
			, 'Microsoft Project Web App'
			, 'Office365 Usage Analytics'
			, 'Power BI JSON Theme Guide'
			, 'Power BI Premium Capacity Metrics'
			, 'Power BI Release Plan'
			, 'Template Apps Exploration Tool'
		)
		[array]$ignoreReports = @(
			'Dashboard Usage Metrics Report'
			, 'Report Usage Metrics Report'
		)
		
		# Get list of workspaces
		$workspaces = Get-PowerBIWorkspace -Scope Organization -All -ErrorAction SilentlyContinue | 
		Where-Object {
			$_.Type -eq 'Workspace' -and
			$_.State -eq 'Active' -and
			$_.Name -notIn $ignoreWorkspaces
		} | Select-Object Name, Id | Sort-Object -Property Name
		
		# If interactive, display a grid view of workspaces and allow the user to select which ones to scan for bare datasets
		$workspaces = $Interactive ? ($workspaces | Out-ConsoleGridView -Title 'Select Workspaces to Scan') : $workspaces
		
		# Declare $bareDatasets array as a concurrent (thread-safe) PSObject
		$bareDatasets = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::New()
		
		# For each workspace, find datasets with no corresponding report and add them to the $bareDatasets array
		$workspaces | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
		
			# Declare local variables
			$workspaceName = $_.Name
			$workspaceId = $_.Id
			$localDatasets = $using:bareDatasets
			
			# Get datasets from the workspace
			$workspaceDatasets = Get-PowerBIDataset -Scope Organization -WorkspaceId $workspaceId -ErrorAction SilentlyContinue |
			Where-Object {
				$_.IsRefreshable -eq $true -and
				$_.Name -notIn $ignoreReports
			} | Select-Object Name, Id, WebUrl, IsRefreshable, @{
				Name = 'WorkspaceName'; Expression = { $workspaceName }
			}, @{
				Name = 'WorkspaceId'; Expression = { $workspaceId }
			} | Sort-Object -Property Name
			
			# Get reports from the workspace
			$workspaceReports = Get-PowerBIReport -Scope Organization -WorkspaceId $workspaceId -ErrorAction SilentlyContinue |
			Where-Object {
				$_.Name -notIn $ignoreReports -and
				$_.WebUrl -notlike '*/rdlreports/*'
			} | Select-Object Name, Id, WebUrl, ReportType, DatasetId, @{
				Name = 'WorkspaceName'; Expression = { $workspaceName }
			}, @{
				Name = 'WorkspaceId'; Expression = { $workspaceId }
			} | Sort-Object -Property Name
			
			# For each dataset, check for any corresponding reports with the same name
			$workspaceDatasets | ForEach-Object {
				$datasetProperties = '' | Select-Object DatasetName, DatasetId, WebUrl, IsRefreshable, WorkspaceName, WorkspaceId
				$datasetName, $datasetId, $datasetWebUrl, $datasetIsRefreshable, $datasetWorkspaceName, $datasetWorkspaceId = $null
				$datasetName = $_.Name
				$datasetId = $_.Id
				$datasetWebUrl = $_.WebUrl
				$datasetIsRefreshable = $_.IsRefreshable
				$datasetWorkspaceName = $_.WorkspaceName
				$datasetWorkspaceId = $_.WorkspaceId
				
				# If no corresponding report is found, add the current dataset to the $bareDatasets array
				if (!($workspaceReports | Where-Object { $_.Name -eq $datasetName -and $_.WorkspaceId -eq $datasetWorkspaceId })) {
					$datasetProperties.DatasetName = $datasetName
					$datasetProperties.DatasetId = $datasetId
					$datasetProperties.WebUrl = $datasetWebUrl
					$datasetProperties.IsRefreshable = $datasetIsRefreshable
					$datasetProperties.WorkspaceName = $datasetWorkspaceName
					$datasetProperties.WorkspaceId = $datasetWorkspaceId
					$localDatasets.Add($datasetProperties)
				}
				
			}
			
		}
		
	}

	return $bareDatasets | Select-Object -Unique -Property DatasetName, DatasetId, WorkspaceName, WorkspaceId

}
