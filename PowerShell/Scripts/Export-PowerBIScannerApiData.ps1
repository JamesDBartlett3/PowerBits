<#
  .SYNOPSIS
    Title: Export-PowerBIScannerApiData
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

  .DESCRIPTION
    This script will get all available data from the Power BI Scanner API and write it to a JSON file.

  .PARAMETER OutFile
    The destination path for the JSON file. Defaults to the Downloads folder.

  .EXAMPLE
    .\Export-PowerBIScannerApiData.ps1 -OutFile "~\Downloads\PowerBIScannerApiData.json"

  .LINK
    https://github.com/JamesDBartlett3/PowerBits

  .LINK
    https://techhub.social/@JamesDBartlett3

  .LINK
    https://datavolume.xyz

  .NOTES
    - Requires the Power BI Management module: https://docs.microsoft.com/en-us/powershell/power-bi/overview?view=powerbi-ps
    - Requires the Power BI Scanner API to be enabled: https://learn.microsoft.com/en-us/power-bi/enterprise/service-admin-metadata-scanning#enabling-enhanced-metadata-scanning
    - Currently only works with workspaces that have been modified in the last 30 days
    - Tenants with a lot of workspaces may not work properly due to API rate limits
#> 

Param(
  [Parameter(Mandatory = $false)]
  [string]$OutFile = '~\Downloads\PowerBIScannerApiData.json'
)

$headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()

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

$getWorkspacesUrl = 'https://api.powerbi.com/v1.0/myorg/admin/workspaces/modified?excludePersonalWorkspaces=True'

# Send a GET request to the API endpoint
$workspaceList = Invoke-RestMethod -Uri $getWorkspacesUrl -Method Get -Headers $headers

# Create an object to hold the workspace IDs
$workspaceIdsObject = [PSCustomObject]@{
  workspaces = @()
}

# Add the workspace IDs to the object
foreach ($w in $workspaceList) {
  $workspaceIdsObject.workspaces += $w.id
}

# Convert the object to JSON
$jsonBody = $workspaceIdsObject | ConvertTo-Json

$startScanUrl = 'https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True'

# Send a POST request to the API endpoint
$startScanResponse = Invoke-RestMethod -Uri $startScanUrl -Method Post -Headers $headers -Body $jsonBody -ContentType 'application/json'

$scanId = $startScanResponse.id

$checkScanUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanStatus/$scanId"

$scanStatus = 'NotStarted'

# Check the scan status every 5 seconds until it's complete
while ($scanStatus -ne 'Succeeded') {
  $checkScanResponse = Invoke-RestMethod -Uri $checkScanUrl -Method Get -Headers $headers
  Start-Sleep -s 5
  $scanStatus = $checkScanResponse.status
  Write-Host "Scan status: $scanStatus"
}

Write-Host 'Scan complete. Getting data...'

$completedScanId = $checkScanResponse.id

$getDataUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$completedScanId"

# Send a GET request to the API endpoint
$getDataResponse = Invoke-RestMethod -Uri $getDataUrl -Method Get -Headers $headers

Write-Host "Writing data to file: $OutFile"

# Write the data to a file
$getDataResponse | ConvertTo-Json -Depth 100 | Out-File -FilePath $OutFile

# Ask user if they want to open the file
$openFile = Read-Host 'Open file? (y/n)'

# Open the file in the default application
if ($openFile) {
  Invoke-Item $OutFile
}