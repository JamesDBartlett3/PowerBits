Param(
  [Parameter(Mandatory)][string]$uri
)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-RestMethod -Method POST -Uri $uri