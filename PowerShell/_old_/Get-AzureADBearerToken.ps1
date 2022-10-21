
# SCRIPT DOWNLOADED FROM: 
# https://community.powerbi.com/t5/Community-Blog/Power-BI-Gateway-Monitoring-amp-Administrating-Part-1/ba-p/750509

# Currently not working

Function Get-AADToken {
  Param(
    [parameter(Mandatory = $true)][string]$Username,
    [parameter(Mandatory = $true)][SecureString]$Password,
    [parameter(Mandatory = $true)][guid]$ClientId,
    [parameter(Mandatory = $true)][string]$path,
    [parameter(Mandatory = $true)][string]$fileName
  )
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$authorityUrl = "https://login.microsoftonline.com/common/oauth2/authorize"

# $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force

  ## load active directory client dll


  $typePath = (Get-ChildItem -Path "C:\Program Files\WindowsPowerShell" `
              -Include "Microsoft.IdentityModel.Clients.ActiveDirectory.dll" `
              -Recurse -Force -ErrorAction SilentlyContinue).FullName |
              Select-Object -First 1

  
  Add-Type -Path $typePath 

  Write-Verbose "Loaded the Microsoft.IdentityModel.Clients.ActiveDirectory.dll"

  Write-Verbose "Using authority: $authorityUrl"
  $authContext = New-Object `
    -TypeName Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext `
    -ArgumentList ($authorityUrl)
  $credential = New-Object `
    -TypeName Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential `
    -ArgumentList ($UserName, $Password)
  
  Write-Verbose "Trying to aquire token for resource: $Resource"
  $authResult = $authContext.AcquireToken("https://analysis.windows.net/powerbi/api", $clientId, $credential)

  Write-Verbose "Authentication Result retrieved for: $($authResult.UserInfo.DisplayableId)"
  
New-Item -path $path -Name $fileName -Value $authResult.AccessToken -ItemType file -force;

return "SuccessFully Writted on the file";


}