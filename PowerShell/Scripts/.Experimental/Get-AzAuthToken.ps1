param(
  [Parameter()]
  [string]$TenantId,
  [Parameter()]
  [string]$AppId,
  [Parameter()]
  [string]$AppSecret
)
process{
  # If Az authentication has not already been established, attempt to authenticate
  if (-not (Get-AzContext) -and $null -eq $AppSecret) {
    try {
      # Convert the app secret (a string) to a SecureString
      $secureAppSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
      # Create a PS credential object
      $psCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $secureAppSecret
      # Attempt to connect to Azure with service principal
      Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $psCred
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