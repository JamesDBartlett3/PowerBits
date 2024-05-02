<#
  .SYNOPSIS
    Get all "Thin" Models (Power BI Semantic Models without a corresponding report) from Power BI Workspaces
  
  .DESCRIPTION
    Get all "Thin" Models (Power BI Semantic Models without a corresponding report) from selected Power BI Workspaces in parallel, and output them to the pipeline.
  
  .PARAMETER ThrottleLimit
    The maximum number of parallel processes to run. Defaults to the number of logical processors in the system.
  
  .PARAMETER Interactive
    If specified, displays a grid view of Workspaces and allows the user to select which ones to scan for Thin Models.
  
  .INPUTS
    This script does not accept pipeline input.
  
  .OUTPUTS
    One or more objects with the following properties:
      - ModelName
      - DatasetId
      - WebUrl
      - IsRefreshable
      - WorkspaceName
      - WorkspaceId
  
  .EXAMPLE
    # Get all Thin Models from all workspaces to which the user has access, one at a time
    .\Get-PowerBIThinModelsFromWorkspaces.ps1 -ThrottleLimit 1
    
  .EXAMPLE
    # Get all Thin Models from all workspaces specified by the user in an interactive prompt, in parallel
    .\Get-PowerBIThinModelsFromWorkspaces.ps1 -Interactive
  
  .NOTES
    This script does NOT require Azure AD app registration, service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the MicrosoftPowerBIMgmt module, if it's not already installed).
    
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
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
  
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits/blob/main/PowerShell/Scripts/Get-PowerBIThinModelsFromWorkspaces.ps1)
  
  .LINK
    [The author's blog](https://datavolume.xyz)
    
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
#>

# PowerShell dependencies
#Requires -PSEdition Core
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
    }
    else {
      Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID authentication dialog...' -ForegroundColor DarkYellow
      Start-Sleep -s 1
      Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
      $headers = Get-PowerBIAccessToken
    }
    if (!$headers) {
			Write-Host '❌ Power BI Access Token not acquired. Exiting...' -ForegroundColor Red
			Exit
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
  $workspaces = Get-PowerBIWorkspace -Scope Organization -All | 
  Where-Object {
    $_.Type -eq 'Workspace' -and
    $_.State -eq 'Active' -and
    $_.Name -notIn $ignoreWorkspaces
  } | Select-Object Name, Id | Sort-Object -Property Name
  
  # If interactive, display a grid view of Workspaces and allow the user to select which ones to scan for Thin Models
  $workspaces = $Interactive ? ($workspaces | Out-ConsoleGridView -Title 'Select Workspaces to Scan') : $workspaces
  
  # Declare $hash as a hashset to store unique Model IDs (prevents duplicates in the output)
  $hash = [System.Collections.Generic.Hashset[guid]]::New()
  
  # For each Workspace, find Datasets with no corresponding report and add them to the $thinModels array
  $workspaces | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
  
    # Declare local variables
    $workspaceName = $_.Name
    $workspaceId = $_.Id
    
    # Get Datasets from the Workspace
    $workspaceModels = Get-PowerBIDataset -Scope Organization -WorkspaceId $workspaceId |
    Where-Object {
      $_.IsRefreshable -eq $true -and
      $_.Name -notIn $ignoreReports
    } | Select-Object Name, Id, WebUrl, IsRefreshable, @{
      Name = 'WorkspaceName'; Expression = { $workspaceName }
    }, @{
      Name = 'WorkspaceId'; Expression = { $workspaceId }
    } | Sort-Object -Property Name
    
    # Get reports from the Workspace
    $workspaceReports = Get-PowerBIReport -Scope Organization -WorkspaceId $workspaceId |
    Where-Object {
      $_.Name -notIn $ignoreReports -and
      $_.WebUrl -notlike '*/rdlreports/*'
    } | Select-Object Name, Id, WebUrl, ReportType, DatasetId, @{
      Name = 'WorkspaceName'; Expression = { $workspaceName }
    }, @{
      Name = 'WorkspaceId'; Expression = { $workspaceId }
    } | Sort-Object -Property Name
    
    # For each Dataset, check for any corresponding reports with the same name
    $workspaceModels | ForEach-Object {
      $datasetProperties = '' | Select-Object ModelName, DatasetId, WebUrl, IsRefreshable, WorkspaceName, WorkspaceId
      $modelName, $datasetId, $datasetWebUrl, $datasetIsRefreshable, $datasetWorkspaceName, $datasetWorkspaceId = $null
      $modelName = $_.Name
      $datasetId = $_.Id
      $datasetWebUrl = $_.WebUrl
      $datasetIsRefreshable = $_.IsRefreshable
      $datasetWorkspaceName = $_.WorkspaceName
      $datasetWorkspaceId = $_.WorkspaceId
      
      # If no corresponding report is found, output the Dataset's properties for processing downstream
      if (!($workspaceReports | Where-Object { $_.Name -eq $modelName -and $_.WorkspaceId -eq $datasetWorkspaceId })) {
        $datasetProperties.ModelName = $modelName
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
  Write-Verbose "Total number of Thin Models: $($hash.Count)"
  
  # Clear the PowerShell session's memory
  [gc]::Collect()
}