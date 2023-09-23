# Title: Get-ArrayInBatches.ps1
# Author: @ruiromano on GitHub
# Source: https://github.com/RuiRomano/pbimonitor
[CmdletBinding()]
Param(
  [Parameter(Mandatory)][array]$array,
  [Parameter(Mandatory)][int]$batchCount,
  [Parameter(Mandatory)][ScriptBlock]$script,
  [Parameter][string]$label = "Get-ArrayInBatches"
)
$skip = 0
$i = 0
do {
  $batchItems = @($array | Select-Object -First $batchCount -Skip $skip)
  if ($batchItems) {
    Write-Host "[$label] Batch: $($skip + $batchCount) / $($array.Count)"
    Invoke-Command -ScriptBlock $script -ArgumentList @($batchItems, $i)
    $skip += $batchCount
  }
  $i++
}
while($batchItems.Count -ne 0 -and $batchItems.Count -ge $batchCount)
