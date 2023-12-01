<#
  .SYNOPSIS
    Title: Export-PowerBIScannerApiData
    Author: James D. Bartlett III

  .DESCRIPTION
    This script will get all available data from the Power BI Scanner API and write it to a JSON file.

  .INPUTS
    - Parameters are currently the only way to pass input to this script
    - Pipeline inputs are not yet supported

  .OUTPUTS
    - A .json file containing all available data from the Power BI Scanner API
    - Pipeline outputs are not yet supported

  .PARAMETER OutFile
    The destination path for the JSON file. Defaults to "~\Downloads\PowerBIScannerApiData_{timestamp}.json"

  .PARAMETER OpenFile
    Specify to open the JSON file in the default application after it's created.

  .EXAMPLE
    .\Export-PowerBIScannerApiData.ps1 -OutFile "C:\temp\MyPowerBIScannerApiData.json"
    # Export data to "C:\temp\MyPowerBIScannerApiData.json"

  .EXAMPLE
    .\Export-PowerBIScannerApiData.ps1 -OpenFile
    # Export data to the default location ("~\Downloads\PowerBIScannerApiData_{timestamp}.json")
    # and open the file in the system's default .json file handler/editor

  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits)

  .LINK
    [The author's blog](https://datavolume.xyz)
    
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)

  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)

  .NOTES
    - Requires the Power BI Management module: https://docs.microsoft.com/en-us/powershell/power-bi/overview?view=powerbi-ps
    - Requires the Power BI Scanner API to be enabled: https://learn.microsoft.com/en-us/power-bi/enterprise/service-admin-metadata-scanning#enabling-enhanced-metadata-scanning
    - Currently only works with workspaces that have been modified in the last 30 days
    - Tenants with a lot of workspaces may not work properly due to API rate limits

    TODO:
      - Add parameters for all available API options

    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>

Param(
  [Parameter(Mandatory = $false)]
  [string]$OutFile = "$HOME\Downloads\PowerBIScannerApiData_$(Get-Date -UFormat '%Y-%m-%d_%H%M').json",
  [Parameter(Mandatory = $false)]
  [switch]$OpenFile
)

$currentDate = Get-Date -UFormat "%Y-%m-%d_%H%M"
$OutFile = $OutFile -replace 'PowerBIScannerApiData.json', "PowerBIScannerApiData_$currentDate.json"
$baseUrl = 'https://api.powerbi.com/v1.0/myorg/admin/workspaces'
$headers = [System.Collections.Generic.Dictionary[[String],[String]]]::New()

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

$modifiedWorkspacesUrl = "$baseUrl/modified?excludePersonalWorkspaces=True"

# Send a GET request to the API endpoint
$workspaceList = Invoke-RestMethod -Uri $modifiedWorkspacesUrl -Method Get -Headers $headers

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

$getInfoUrl = "$baseUrl/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True"

# Send a POST request to the API endpoint
$startScanResponse = Invoke-RestMethod -Uri $getInfoUrl -Method Post -Headers $headers -Body $jsonBody -ContentType 'application/json'

$scanId = $startScanResponse.id

$scanStatusUrl = "$baseUrl/scanStatus/$scanId"

$scanStatus = ''

# Check the scan status every 5 seconds until it's complete
while ($scanStatus -ne 'Succeeded') {
  $checkScanResponse = Invoke-RestMethod -Uri $scanStatusUrl -Method Get -Headers $headers
  Start-Sleep -s 5
  $scanStatus = $checkScanResponse.status
  Write-Host "Scan status: $scanStatus"
}

Write-Host 'Scan complete. Getting data...'

$completedScanId = $checkScanResponse.id

$scanResultUrl = "$baseUrl/scanResult/$completedScanId"

# Send a GET request to the API endpoint
$getDataResponse = Invoke-RestMethod -Uri $scanResultUrl -Method Get -Headers $headers

Write-Host "Writing data to file: $OutFile"

# Write the data to a file
$getDataResponse | ConvertTo-Json -Depth 100 | Out-File -FilePath $OutFile

# Open the file in the default application if user passed the -OpenFile switch
if ($OpenFile) {
  Invoke-Item $OutFile
}