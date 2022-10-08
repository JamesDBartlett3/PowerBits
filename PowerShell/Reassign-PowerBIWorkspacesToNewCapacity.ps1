
# Bearer Token should include the word 'Bearer' and a space, e.g. Bearer <token>
$bearerToken = ''
$oldCapacityId = ''
$newCapacityId = ''


$baseUrl = 'wabi-us-north-central-h-primary-redirect.analysis.windows.net'
$requestUri = 'https://' + $baseUrl + "/capacities/$oldCapacityId/workspaces"
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

$result = Invoke-WebRequest -UseBasicParsing -Uri $requestUri -WebSession $session -Headers @{
  'method'               = 'GET'
  'authority'            = $baseUrl
  'path'                 = "/capacities/$oldCapacityId/workspaces"
  'authorization'        = $bearerToken
} | ConvertFrom-Json

$result.workspaceObjectId | ForEach-Object {
  Invoke-WebRequest -Method POST -Uri "https://api.powerbi.com/v1.0/myorg/groups/$_/AssignToCapacity" -Body @{
    'capacityId' = $newCapacityId
    } -Headers @{
      'authorization' = $bearerToken
    }
}