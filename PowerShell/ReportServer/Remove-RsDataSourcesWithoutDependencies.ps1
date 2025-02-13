#***********************************************************************************************************************************************
# FileName:             Remove-RsDataSourcesWithoutDependencies.ps1
# Date:                 2024-06-11
# Original Author:      Hugh Scott (https://stackoverflow.com/a/29827177)
# Modified by:          James Bartlett
#
# Description:
# This script finds data sources with no dependencies in SSRS/PBIRS and removes them.
#
# Parameters:
#   $ServerName               - Name of the server on which the reportserver is running (e.g. reportserver.example.com)
#   $ReportServerInstanceName - [Optional] Name of the reportserver instance (default: ReportServer)
#   $PortNumber               - [Optional] Port number on which the reportserver is running (default: 443)
#   $OutFile                  - [Optional] Path to logfile which will list all actions taken by the script (default: a txt file in the same dir)
#   $WhatIf                   - [Optional] Switch parameter to list out reports that would have been deleted (instead of actually deleting them)
#***********************************************************************************************************************************************

[CmdletBinding()]
Param(
  [Parameter(Mandatory)][string]$ServerName,
  [Parameter()][string]$ReportServerInstanceName = "ReportServer",
  [Parameter()][int]$PortNumber = 443,
  [Parameter()][string]$OutFile = (Join-Path -Path $PSScriptRoot -ChildPath "$((Get-Date).ToString('yyyyMMdd-HHmm'))_$($ServerName)_DataSourcesWithoutDependencies.txt"),
  [Parameter()][switch]$WhatIf
)

#Requires -Version 5.1

$url = "https://$($ServerName):$($PortNumber)/$($ReportServerInstanceName)/ReportService2010.asmx?wsdl"
$ssrs = New-WebServiceProxy -uri $url -UseDefaultCredential -Namespace "ReportingWebService"

# Connect to ReportingWebService, grab all data sources, and delete those without dependencies
$items = $ssrs.ListChildren("/", $true) | Where-Object {$_.typename -eq "DataSource"}
Set-Content $OutFile "Items $(if($WhatIf) {'which would be '})deleted:"
foreach($item in $items) {

  $dependencies = $ssrs.ListDependentItems($item.Path)
  $dependentReports = $dependencies.Count

  if($dependentReports -eq 0){
    [string]$itemName = $item.Path
    if($WhatIf){
      Write-Host "Item $itemName would be deleted."
      Add-Content $OutFile $itemName
    } else {
      try {
        $ssrs.DeleteItem($item.Path)
        Write-Host "Item $itemName deleted."
        Add-Content $OutFile $itemName
      } catch [System.Exception] {
        $Msg = $_.Exception.Message
        Write-Host $itemName $Msg
        Add-Content $itemName $Msg
      }
    }
  }
}