# WARNING: This script utilizes an undocumented internal API, and is thus not supported by Microsoft. Use at your own considerable risk.
try {
  Get-PowerBIAccessToken
} catch {
  Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID authentication dialog...'
  Start-Sleep -s 1
  Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
  Get-PowerBIAccessToken
} finally {
  Write-Host '🔑 Power BI Access Token acquired.'
  $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
  $session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36 Edg/110.0.1587.63"
  $myRegionUrl = "wabi-us-north-central-h-primary-redirect.analysis.windows.net" # Replace with your own region's URL
  $myBearerToken = (Get-PowerBIAccessToken)["Authorization"] # Add your own bearer token here
  (Invoke-WebRequest -UseBasicParsing -Uri "https://$myRegionUrl/metadata/gallery/SharedDatasets" `
  -WebSession $session -Headers @{
  "authority"="$myRegionUrl"
    "method"="GET"
    "path"="/metadata/gallery/SharedDatasets"
    "scheme"="https"
    "accept"="application/json, text/plain, */*"
    "accept-encoding"="gzip, deflate, br"
    "accept-language"="en-US,en;q=0.9"
    "Authorization"="$myBearerToken"
    "cache-control"="no-cache"
    "dnt"="1"
    "origin"="https://app.powerbi.com"
    "pragma"="no-cache"
    "referer"="https://app.powerbi.com/"
    "sec-ch-ua"="`"Chromium`";v=`"110`", `"Not A(Brand`";v=`"24`", `"Microsoft Edge`";v=`"110`""
    "sec-ch-ua-mobile"="?0"
    "sec-ch-ua-platform"="`"Windows`""
    "sec-fetch-dest"="empty"
    "sec-fetch-mode"="cors"
    "sec-fetch-site"="cross-site"
    "x-powerbi-hostenv"="Power BI Web"
  }).Content | ConvertFrom-Json -AsHashTable |
    ForEach-Object {$_ | Select-Object WorkspaceName -ExpandProperty model} |
    Select-Object @{l="Workspace";e="workspaceName"},@{l="Dataset";e="displayName"}, @{l="Size (MB)"; e="sizeInMBs"} |
    Sort-Object -Property Workspace, Dataset
}