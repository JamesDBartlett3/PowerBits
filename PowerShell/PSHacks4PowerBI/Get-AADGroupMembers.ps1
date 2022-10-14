<# 

.SYNOPSIS
  Function: Get-AADGroupMembers
  Author: @JamesDBartlett3 (James D. Bartlett III)

.DESCRIPTION
  - Gets all members of an Azure AD group

.PARAMETERS
  - GroupNameSearchString: String to search for in group name

.TODO
  - 

.EXAMPLE
  Get-AADGroupMembers -GroupNameSearchString "Power BI"

#>

Function Get-AADGroupMembers {
  #Requires -PSEdition Core
  #Requires -Modules Az.Resources
  Param(
    [parameter(Mandatory = $true)][string]$GroupNameSearchString
  )
  Get-AzADGroup -SearchString $GroupNameSearchString | 
  ForEach-Object {
    $group = $_;
    Get-AzADGroupMember -GroupObjectId $group.Id | 
    Select-Object DisplayName, @{
      l = "GroupName";
      e = { $group.DisplayName }
    }
  }

}