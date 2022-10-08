
# Title: Get-UserDatasets
# Author: @JamesDBartlett3
# Parameters: $userEmail
# Returns: "Id" column with DatasetId for all Power BI datasets marked as "configured by" the given user
# Usage: Get-UserDatasets user@domain.tld

#Requires -Modules MicrosoftPowerBIMgmt.Profile

Param(
  [string]$userEmail
)

$hadToLogin = $false
$ignoreReports = "Report Usage Metrics Report", "Dashboard Usage Metrics Report"

try {
  Get-PowerBIAccessToken | Out-Null
}
catch {
  $hadToLogin = $true
  Connect-PowerBIServiceAccount | Out-Null
}
finally{
  $result = Get-PowerBIDataset -Scope Organization |
    Where-Object -Property ConfiguredBy -eq $userEmail |
    Where-Object -Property Name -NotIn $ignoreReports |
    Select-Object -Property Id, Name
}

if($hadToLogin) {
  Disconnect-PowerBIServiceAccount
}

return $result
