<###############################################################/
Author: James D. Bartlett III (@jamesdbartlett3@techhub.social)
Date: 2023-11-08
Purpose: Backup Report Server items to a folder
Requirements:
  - ReportingServicesTools module installed in Windows PowerShell
/###############################################################>

#Requires -Module ReportingServicesTools

param(
  [Parameter()][string]$TargetDirectory = (Join-Path -Path $env:TEMP -ChildPath ReportServer_Backup)
  , [Parameter(Mandatory)][string]$ServerName
  , [Parameter()][int]$ServerPort = 443
)

# If PowerShell version is greater than 5, import ReportingServicesTools module with -UseWindowsPowerShell parameter
if ($PSVersionTable.PSVersion.Major -gt 5) {
  Remove-Module ReportingServicesTools | Out-Null
  try {
    Import-Module ReportingServicesTools -UseWindowsPowerShell
  } catch {
    Write-Error "Unable to import ReportingServicesTools module with -UseWindowsPowerShell parameter. Please install ReportingServicesTools module in Windows PowerShell and try again."
    Write-Host "Try: powershell.exe -NoProfile -ExecutionPolicy Bypass -Command 'Install-Module ReportingServicesTools -Scope CurrentUser'"
    exit
  }
}

# Declare Report Server URI
$sourceRsUri = "https://$($ServerName):$($ServerPort)/ReportServer/"

# Declare Proxy
$proxy = New-RsWebServiceProxy -ReportServerUri $sourceRsUri

# Get all catalog items NOT in subfolders of "/Users Folders"
$proxy.ListChildren("/", $false) | Where-Object { $_.TypeName -eq "Folder" -and $_.Path -notlike "/Users Folders*" } | ForEach-Object {
  Write-Host "Processing folder $($_.Path)..."
  $subfolder = (Join-Path -Path $TargetDirectory -ChildPath $_.Path)
  if (!(Test-Path -Path $subfolder)) {
    New-Item -Path $subfolder -ItemType Directory | Out-Null
  }
  Out-RsFolderContent -Proxy $proxy -RsFolder $_.Path -Recurse -Destination $subfolder
}