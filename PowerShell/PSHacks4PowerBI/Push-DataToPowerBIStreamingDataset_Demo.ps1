<# 
  .SYNOPSIS
    Function: Push-DataToPowerBIStreamingDataset_Demo
    Author: @JamesDBartlett3 (James D. Bartlett III)

  .DESCRIPTION
    - Power BI Push Datasets Demo

  .PARAMETERS
  - 

  .SETUP
    - http://blogs.lobsterpot.com.au/2020/07/16/getting-started-with-power-bi-push-datasets-via-rest-apis

  .TODO
    - Convert to function
    - Parameterize (endpoint, categories, min/max random value, etc.)
    - Determine if token is needed and implement if so

#>

#Requires -PSEdition Core

$endpoint = "" # Endpoint URL

while($true) {

  $timeStamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mmZ')
  $randomNumber0 = (Get-Random -Maximum 100) / 100
  $randomNumber1 = Get-Random -Maximum 500
  $randomString = "yes", "no", "maybe", "meh" | Get-Random
  $randomColor = "red", "yellow", "blue" | Get-Random

  $payload = $null

  $payload = @{
    "someNumber0" = $randomNumber0
    "someNumber1" = $randomNumber1
    "someText" = $randomString
    "someColor" = $randomColor
    "timeStamp" = $timeStamp
  }

  Invoke-RestMethod -Method Post -Uri "$endpoint" -Body (ConvertTo-Json @($payload)) | Out-Null

  Start-Sleep -Seconds 7.5 | Out-Null

}
