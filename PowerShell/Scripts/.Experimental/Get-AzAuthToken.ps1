param(
  [Parameter()]
  [string]$tenantId,
  [Parameter()]
  [string]$appId,
  [Parameter()]
  [string]$appSecret
)
process{
  # If Az authentication has not already been established, attempt to authenticate
  if (-not (Get-AzContext) -and $null -eq $appSecret) {
    try {
      # Convert the app secret (a string) to a SecureString
      $secureAppSecret = ConvertTo-SecureString -String $appSecret -AsPlainText -Force
      # Create a PS credential object
      $psCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $appId, $secureAppSecret
      # Attempt to connect to Azure with service principal
      Connect-AzAccount -ServicePrincipal -Tenant $tenantId -Credential $psCred
    }
    catch {
      # If service principal authentication fails, fall back to interactive user authentication
      Write-Output "Service principal authentication failed. Falling back to interactive user authentication."
      Connect-AzAccount
    }
  }
  # Return the access token
  return (Get-AzAccessToken)
}