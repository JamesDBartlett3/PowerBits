#Requires -PSEdition Core
Function Get-ArrayInBatches {
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
}
Function Wait-On429Error {
  # Title: Wait-On429Error.ps1 
  # Author: @ruiromano on GitHub
  # Source: https://github.com/RuiRomano/pbimonitor
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)][ScriptBlock]$script,
    [Parameter][int]$sleepSeconds = 3601,
    [Parameter][int]$tentatives = $null
  )
  try {
    Invoke-Command -ScriptBlock $script
  } catch {
    $ex = $_.Exception
    $errorText = $ex.ToString()
    if ($errorText -like '*HttpRequestException*' -and $errorText -like '*429*') {
      Write-Host "'429 (Too Many Requests)' Error - Sleeping for $sleepSeconds seconds before trying again..." -ForegroundColor Yellow
      if ($tentatives) {$tentatives = $tentatives - 1}
      if ($tentatives -eq 0) {
        throw '[Wait-On429Error] Max tentatives reached!'
      } else {
        Start-Sleep -Seconds $sleepSeconds
        Wait-On429Error -script $script -sleepSeconds $sleepSeconds -tentatives $tentatives
      }
    } else {
      throw
    }
  }
}
