Login-PowerBIServiceAccount

Get-PowerBIWorkspace -Scope Organization -All | 
  Where-Object -Property "Type" -Eq "Workspace" | 
  Out-GridView -PassThru