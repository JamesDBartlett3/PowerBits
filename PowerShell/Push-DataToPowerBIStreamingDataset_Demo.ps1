<# 

Title: Push-DataToPowerBIStreamingDataset_Demo.ps1
Description: Power BI Push Datasets Demo
Author: @jamesdbartlett3

Setup instructions: 
http://blogs.lobsterpot.com.au/2020/07/16/getting-started-with-power-bi-push-datasets-via-rest-apis

#>

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
