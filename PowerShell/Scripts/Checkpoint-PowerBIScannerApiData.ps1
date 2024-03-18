<#
  .SYNOPSIS
    Exports latest data from the Power BI Scanner API to a JSON file.
  
  .DESCRIPTION
    This script will get the latest available data from the Power BI Scanner API and write it to a .json file.
    By default, the file will be saved in the user's Downloads folder with a timestamp in the filename.
    The user can specify a custom path for the file to be saved by passing the -OutFile parameter.
    The user can also specify to open the file in the default application after it's created by passing the -OpenFile switch.
  
  .PARAMETER OutFile
    The destination path for the JSON file. Defaults to "~\Downloads\PowerBIScannerApiData_{timestamp}.json"
  
  .PARAMETER OpenFile
    Specify to open the JSON file in the default application after it's created.

  .PARAMETER BatchSize
    The number of workspaces to scan in each batch. Valid values are 1-100. Defaults to 100.

  .PARAMETER Organization
    The name of the organization for the REST API baseURL. Defaults to 'myorg'.
  
  .INPUTS
    - Parameters are currently the only way to pass input to this script
    - Pipeline inputs are not yet supported
  
  .OUTPUTS
    - A .json file containing all available data from the Power BI Scanner API
    - Pipeline outputs are not yet supported
  
  .EXAMPLE
    # Export data to "C:\temp\MyPowerBIScannerApiData.json"
    .\Checkpoint-PowerBIScannerApiData.ps1 -OutFile "C:\temp\MyPowerBIScannerApiData.json"
  
  .EXAMPLE
    # Export data to the default location ("~\Downloads\PowerBIScannerApiData_{timestamp}.json")
    # and open the file in the system's default .json file handler/editor
    .\Checkpoint-PowerBIScannerApiData.ps1 -OpenFile
  
  .NOTES
    - Requires the Power BI Management module: https://docs.microsoft.com/en-us/powershell/power-bi/overview?view=powerbi-ps
    - Requires the Power BI Scanner API to be enabled: https://learn.microsoft.com/en-us/power-bi/enterprise/service-admin-metadata-scanning#enabling-enhanced-metadata-scanning
    - Currently only works with Power BI workspaces that have been modified in the last 30 days
    - May not work properly in Power BI tenants with a lot of active workspaces (due to API rate limits)
    
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
  
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits/blob/main/PowerShell/Scripts/Checkpoint-PowerBIScannerApiData.ps1)
  
  .LINK
    [The author's blog](https://datavolume.xyz)
    
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
#>

Param(
  [Parameter(Mandatory = $false)]
  [string]$OutFile = "$HOME\Downloads\PowerBIScannerApiData_$(Get-Date -UFormat '%Y-%m-%d_%H%M').json",
  [Parameter(Mandatory = $false)]
  [switch]$OpenFile,
  [Parameter(Mandatory = $false)][ValidateRange(1, 100)]
  [int]$BatchSize = 100,
  [Parameter(Mandatory = $false)]
  [string]$Organization = 'myorg'
)

#Requires -Modules MicrosoftPowerBIMgmt

# Declare starting variables
[string]$baseUrl = "https://api.powerbi.com/v1.0/$organization/admin/workspaces"
[string]$modifiedWorkspacesUrl = "$baseUrl/modified?excludePersonalWorkspaces=True"
[string]$getInfoUrl = "$baseUrl/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True"
$headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()
$scanResults = [PSCustomObject]@{
  workspaces                       = @()
  datasourceInstances              = @()
  misconfiguredDatasourceInstances = @()
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

# Send a GET request to the modified workspaces endpoint
$workspaceList = Invoke-RestMethod -Uri $modifiedWorkspacesUrl -Method Get -Headers $headers

# Create an object to hold the workspace IDs
$workspaceIdsObject = [PSCustomObject]@{
  workspaces = @()
}

# Add the workspace IDs to the object
foreach ($w in $workspaceList) {
  $workspaceIdsObject.workspaces += $w.id
}

# Get the number of workspaces
[int]$workspaceCount = $workspaceIdsObject.workspaces.Count

# Declare a variable to hold the workspace suffix (singular or plural)
[string]$workspaceSuffix = if ($workspaceCount -eq 1) {''} else {'s'}

# If no workspaces were found, exit the script
if ($workspaceCount -eq 0) {
  Write-Host 'No workspaces found. Exiting...'
  Exit
}

# Calculate the number of batches to run based on the number of workspaces and the batch size
[int]$batchesToRun = [Math]::Ceiling($workspaceCount / $batchSize)

# Declare a variable to hold the batch suffix (singular or plural)
[string]$batchSuffx = if ($batchesToRun -eq 1) {''} else {'es'}

Write-Host "Found $($workspaceCount) workspace$workspaceSuffix. Running $batchesToRun batch$batchSuffx of $batchSize..."
Write-Host '----------------------------------'

# Loop through the workspaces in batches of $batchSize, get the scanner API data for each batch, and add it to the $scanResults object
for ($i = 0; $i -lt $batchesToRun; $i++) {
  $batchNum = $i + 1
  Write-Host "Running batch $batchNum of $batchesToRun..."
  $batchStart = $i * $batchSize
  $batchEnd = $batchStart + $batchSize
  $batch = $workspaceIdsObject.workspaces[$batchStart..$batchEnd]
  $batchObject = [PSCustomObject]@{
    workspaces = $batch
  }
  $jsonBody = $batchObject | ConvertTo-Json
  # Send a POST request to the API endpoint
  $startScanResponse = Invoke-RestMethod -Uri $getInfoUrl -Method Post -Headers $headers -Body $jsonBody -ContentType 'application/json'
  $scanId = $startScanResponse.id
  $scanStatusUrl = "$baseUrl/scanStatus/$scanId"
  $scanStatus = ''
  # Check the scan status every 5 seconds until it's complete
  while ($scanStatus -ne 'Succeeded') {
    if ($scanStatus -eq 'Running') {
      Start-Sleep -s 5
    }
    $checkScanResponse = Invoke-RestMethod -Uri $scanStatusUrl -Method Get -Headers $headers
    $scanStatus = $checkScanResponse.status
    Write-Host "Batch $batchNum status: $scanStatus"
  }
  Write-Host "Batch $batchNum complete. Getting data..."
  Write-Host '----------------------------------'

  $completedScanId = $checkScanResponse.id

  $batchResultUrl = "$baseUrl/scanResult/$completedScanId"
  
  # Send a GET request to the API endpoint
  $batchResult = Invoke-RestMethod -Uri $batchResultUrl -Method Get -Headers $headers

  # Add the result to the $scanResults object
  $scanResults.workspaces += $batchResult.workspaces
  $scanResults.datasourceInstances += $batchResult.datasourceInstances
  $scanResults.misconfiguredDatasourceInstances += $batchResult.misconfiguredDatasourceInstances

}

Write-Host "Finished scanning $workspaceCount workspace$workspaceSuffix ($batchesToRun batch$batchSuffx of $BatchSize)." 
Write-Host "Writing data to file: $OutFile"

# Write the data to $OutFile
$scanResults | ConvertTo-Json -Depth 100 | Out-File -FilePath $OutFile

# Open the file in the default application if user passed the -OpenFile switch
if ($OpenFile) {
  Write-Host 'Opening file...'
  Invoke-Item $OutFile
}