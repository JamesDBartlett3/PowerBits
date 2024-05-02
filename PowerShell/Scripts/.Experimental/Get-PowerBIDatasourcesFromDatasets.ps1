<#
.SYNOPSIS
  Get the datasources for all datasets in a Power BI tenant.
.DESCRIPTION
  Get the datasources for all datasets in a Power BI tenant, optionally filtered by datasource type, server name, and/or database name.
.PARAMETER DatasourceType
  The datasource type to filter on.
.PARAMETER DatasourceServer
  The datasource server to filter on.
.PARAMETER DatasourceDatabase
  The datasource database to filter on.
.PARAMETER BatchStart
  The index of the first dataset to return.
.PARAMETER BatchEnd 
  The index of the last dataset to return.
.EXAMPLE
  # Get the datasources for all datasets in a Power BI tenant, filtered by datasource type, server, and database.
  .\Get-PowerBIDatasourcesFromDatasets.ps1 -DatasourceType 'sql' -DatasourceServer 'myserver.database.windows.net' -DatasourceDatabase 'mydatabase'
.NOTES
  TODO: Refactor to use the Scanner API instead of the Admin API.
#>

[CmdletBinding()]
Param (
  [Parameter()][string]$DatasourceType,
  [Parameter()][string]$DatasourceServer,
  [Parameter()][string]$DatasourceDatabase,
  [Parameter()][int]$BatchStart,
  [Parameter()][int]$BatchEnd
)

#Requires -Modules MicrosoftPowerBIMgmt -PSEdition Core

begin {
  $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()
  $pbiApiBaseUrl = 'https://api.powerbi.com/v1.0/myorg'
  $pbiDatasetsAdminEndpoint = 'admin/datasets'
}

process {

  try {
    $headers = Get-PowerBIAccessToken
  }

  catch {
    if ($servicePrincipalId) {
      $headers = Connect-PowerBIServiceAccount -ServicePrincipal -Tenant $servicePrincipalTenantId -Credential $credential
    } else {
      Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID (f.k.a. Azure Active Directory) authentication dialog...'
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

  $datasets = (Invoke-RestMethod -Uri "$pbiApiBaseUrl/$pbiDatasetsAdminEndpoint" -Headers $headers -Method Get).value

  $datasets = $datasets | Where-Object { $_.IsRefreshable -eq $true -and $_.contentProviderType -eq 'PbixInImportMode' }

  [PSCustomObject]$datasources = foreach ($dataset in $datasets[$BatchStart..$BatchEnd]) {

    $datasetId = $dataset.id
    $datasetName = $dataset.name
    $workspaceId = $dataset.workspaceId

    # Get the datasources for the current dataset
    $datasetDatasources = (Invoke-RestMethod -Uri "$pbiApiBaseUrl/$pbiDatasetsAdminEndpoint/$($dataset.id)/datasources" -Headers $headers -Method Get).value

    # Return the dataset and its datasources
    [PSCustomObject]@{
      DatasetId = $datasetId
      DatasetName = $datasetName
      WorkspaceId = $workspaceId
      Datasources = $datasetDatasources
    }

  }

  $result = $datasources | ForEach-Object {
    $datasetId = $_.DatasetId
    $datasetName = $_.DatasetName
    $workspaceId = $_.WorkspaceId
    $dsType = $_.Datasources.datasourceType
    $dsServer = $_.Datasources.connectionDetails.server
    $dsDatabase = $_.Datasources.connectionDetails.database
    $dsId = $_.Datasources.datasourceId
    $dsGatewayId = $_.Datasources.gatewayId

    [PSCustomObject]@{
      DatasetId = $datasetId
      DatasetName = $datasetName
      WorkspaceId = $workspaceId
      DatasourceType = $dsType -join ','
      DatasourceServer = $dsServer -join ','
      DatasourceDatabase = $dsDatabase -join ','
      DatasourceId = $dsId -join ','
      DatasourceGatewayId = $dsGatewayId -join ','
    }

  }

  # Filter results using the parameters
  if ($DatasourceType) {
    $result = $result | Where-Object { $_.DatasourceType -icontains $DatasourceType }
  }
  if ($DatasourceServer) {
    $result = $result | Where-Object { $_.DatasourceServer -icontains $DatasourceServer }
  }
  if ($DatasourceDatabase) {
    $result = $result | Where-Object { $_.DatasourceDatabase -icontains $DatasourceDatabase }
  }
  
  return $result

}