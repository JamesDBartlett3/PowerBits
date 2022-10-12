## NOT WORKING CURRENTLY ##
## Potential fix: https://github.com/microsoft/powerbi-powershell/issues/140 ##

$identifier = "Vitals@dmu.edu" # Identifier Options: <user_email_address>, <group_name>, <group_guid>

$testing = $false

$pbiToken = $null
$pbiToken = Get-PowerBIAccessToken
if(!$pbiToken){
  Login-PowerBIServiceAccount | Out-Null
}

$workspaces = Get-PowerBIWorkspace -Scope Organization -All | 
  Where-Object {
    $_.State -EQ "Active" -AND
    $_.IsOrphaned -EQ $False -AND
    $_.Type -EQ "Workspace"
  } | Select-Object -Property Id, Name | Sort-Object -Property Name |
    Out-GridView -PassThru -Title "Select Workspaces (Ctrl+Click or Shift+Click to Multi-Select)"

ForEach ($w in $workspaces) {
  $name = $w.Name
  $Params = @{
    Scope = "Organization"
    Id = $w.Id
    UserPrincipalName = $identifier
    WarningAction = "SilentlyContinue"
  }
  Write-Output "Removing `e[38;2;0;255;0m$identifier`e[0m from workspace `e[38;2;255;0;0m$name`e[0m "
  Remove-PowerBIWorkspaceUser @Params | Out-Null
  Write-Output "Waiting 18 seconds to avoid hitting the API limit (200 req/hr)..."
  Start-Sleep 18
}

if(-not $testing) {
  Logout-PowerBIServiceAccount | Out-Null
}
