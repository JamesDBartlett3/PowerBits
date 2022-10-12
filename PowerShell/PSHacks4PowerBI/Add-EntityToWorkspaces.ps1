## NOT WORKING CURRENTLY ##
## Potential fix: https://github.com/microsoft/powerbi-powershell/issues/140 ##

$identifier = "Vitals@dmu.edu" # Identifier Options: <user_email_address>, <group_name>, <group_guid>
$principalType = "Group" # PrincipalType Options: App, Group, User
$accessRight = "Admin" # AccessRight Options: Member, Admin, Contributor, Viewer

try {

  Get-PowerBIAccessToken

} catch {

  Write-Output "Power BI Access Token required. Launching authentication dialog..."
  Start-Sleep -s 2
  Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null

}

finally {

  $workspaces = Get-PowerBIWorkspace -Scope Organization -All | 
  Where-Object {
    $_.State -EQ "Active" -AND
    $_.Type -EQ "Workspace"
    } | Select-Object -Property Id, Name, State  | Sort-Object -Property Name |
    Out-GridView -PassThru -Title "Select Workspaces (Ctrl+Click or Shift+Click to Multi-Select)"

  $objectId = $identifier

  if ($workspaces) {

    if($principalType -EQ "Group"){
      $aadToken = Get-AzureADCurrentSessionInfo
      if(!$aadToken){
        if($PSVersionTable.PSEdition -EQ "Core"){
          Remove-Module AzureAD
          Import-Module AzureAD -UseWindowsPowerShell -WarningAction SilentlyContinue
        } else {
          Import-Module AzureAD
        }
        Connect-AzureAD | Out-Null
      }
      $objectId = Get-AzureADGroup -SearchString $identifier |
        ForEach-Object {$_.ObjectId} | Select-Object -First 1
    }

    ForEach ($w in $workspaces) {
      $name = $w.Name
      $Params = @{
        Scope = "Organization"
        Id = $w.Id
        Identifier = $objectId
        PrincipalType = $principalType
        AccessRight = $accessRight
        WarningAction = "SilentlyContinue"
      }
      Write-Output "Adding `e[38;2;0;255;0m$identifier`e[0m to workspace `e[38;2;255;0;0m$name`e[0m"
      Add-PowerBIWorkspaceUser @Params | Out-Null
      Write-Output "Waiting 18 seconds to avoid hitting the API limit (200 req/hr)..."
      Start-Sleep 18
    }

  } else {

    Write-Output "No workspaces specified. Exiting..."
  
  }

}
