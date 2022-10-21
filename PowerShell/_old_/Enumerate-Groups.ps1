Get-PowerBIWorkspace -Scope Organization -All |
  Where-Object -Property "Type" -Eq "Group" |
  Out-GridView -PassThru 