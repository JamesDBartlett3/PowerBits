<#
  .SYNOPSIS
    Backup Report Server (SSRS/PBIRS) items to a folder
  .DESCRIPTION
    This script will create a folder for each Report Server folder and save the contents of each folder to a subfolder of the target directory. 
    If you are using PowerShell 5 or greater, this script will automatically re-import the ReportingServicesTools module with the -UseWindowsPowerShell parameter.
  .INPUTS
    - Parameters are currently the only way to pass input to this script
    - Pipeline inputs are not yet supported
  .PARAMETER TargetDirectory (Optional)
    The target directory to save the backup to (Defaults to $env:TEMP\{ServerNameParameterValue}_ReportServer_Backup_{CurrentDate}_{CurrentTime})
    So, if you run this script with the default TargetDirectory parameter value on a server named "MyReportServer" on 2023-11-08 at 10:30:00, the target directory will be:
    C:\Users\{YourUserName}\AppData\Local\Temp\MyReportServer_ReportServer_Backup_20231108_103000
  .PARAMETER ServerName (Required)
    The name of the Report Server
  .PARAMETER ServerPort (Optional)
    The port of the Report Server (Defaults to 443)
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer"
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer" -ServerPort 80
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer" -TargetDirectory "C:\Temp\ReportServer_Backup"
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits)
  .LINK
    [The author's blog](https://datavolume.xyz)
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
  .NOTES
    Version:        1.0
    Author:         James D. Bartlett III (@jamesdbartlett3@techhub.social)
    Creation Date:  2023-11-08
    Acknowledgements:
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>

#Requires -Module ReportingServicesTools

param(
  [Parameter()][string]$TargetDirectory = (Join-Path -Path $env:TEMP -ChildPath ReportServer_Backup)
  , [Parameter(Mandatory)][string]$ServerName
  , [Parameter()][int]$ServerPort = 443
)

# If user did not specify a target directory, add the ServerName and current datetime to the default target directory
if($TargetDirectory -eq (Join-Path -Path $env:TEMP -ChildPath ReportServer_Backup)) {
  $TargetDirectory = $TargetDirectory.Replace((Split-Path $TargetDirectory -Leaf), "$($ServerName)_ReportServer_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')")
}

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