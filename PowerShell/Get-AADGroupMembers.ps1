
Param(
  [parameter(Mandatory = $true)][string]$GroupNameSearchString
)

Get-AzADGroup -SearchString $GroupNameSearchString | 
ForEach-Object {
  $group = $_;
  Get-AzADGroupMember -GroupObjectId $group.Id | 
  Select-Object DisplayName, @{
    l="GroupName";
    e={$group.DisplayName}
  }
}