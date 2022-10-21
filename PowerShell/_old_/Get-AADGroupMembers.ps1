<# 
  .SYNOPSIS
    Function: Get-AADGroupMembers
    Author: @JamesDBartlett3 (James D. Bartlett III)

  .DESCRIPTION
    - Gets all members of an Azure AD group

  .PARAMETER GroupName
    String to search for in group name

  .EXAMPLE
    Get-AADGroupMembers -GroupName "Power BI"

  .NOTES
    TODO
      - Testing
      - Add usage, help, and examples.
#>

Function Get-AADGroupMembers {
  #Requires -PSEdition Core
  #Requires -Modules Az.Resources
  Param(
    [parameter(Mandatory = $true)][string]$GroupName
  )
  Get-AzADGroup -SearchString $GroupName | 
  ForEach-Object {
    $group = $_;
    Get-AzADGroupMember -GroupObjectId $group.Id | 
    Select-Object DisplayName, @{
      l = "GroupName";
      e = { $group.DisplayName }
    }
  }
}