begin {
  $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()
  $pbiApiBaseUrl = 'https://api.powerbi.com/v1.0/myorg'
  $pbiDatasetsAdminEndpoint = '/admin/datasets'
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

  $datasets = (Invoke-RestMethod -Uri "$pbiApiBaseUrl$pbiDatasetsAdminEndpoint" -Headers $headers -Method Get).value

  $datasets = $datasets | Where-Object { $_.IsRefreshable -eq $true -and $_.contentProviderType -eq 'PbixInImportMode' }
  

  [PSCustomObject]$datasources = foreach ($dataset in $datasets[0..50]) {

    $datasetId = $dataset.id
    $datasetName = $dataset.name
    $workspaceId = $dataset.workspaceId

    # Get the datasources for the current dataset
    $datasetDatasources = (Invoke-RestMethod -Uri "$pbiApiBaseUrl$pbiDatasetsAdminEndpoint/$($dataset.id)/datasources" -Headers $headers -Method Get).value

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
    $datasourceType = $_.Datasources.datasourceType
    $datasourceServer = $_.Datasources.connectionDetails.server
    $datasourceDatabase = $_.Datasources.connectionDetails.database
    $datasourceId = $_.Datasources.datasourceId
    $datasourceGatewayId = $_.Datasources.gatewayId

    [PSCustomObject]@{
      DatasetId = $datasetId
      DatasetName = $datasetName
      WorkspaceId = $workspaceId
      DatasourceType = $datasourceType -join ','
      DatasourceServer = $datasourceServer -join ','
      DatasourceDatabase = $datasourceDatabase -join ','
      DatasourceId = $datasourceId -join ','
      DatasourceGatewayId = $datasourceGatewayId -join ','
    }

  }
  
  $result | Where-Object { $_.DatasourceType -contains 'Sql' } | Format-Table -AutoSize

}