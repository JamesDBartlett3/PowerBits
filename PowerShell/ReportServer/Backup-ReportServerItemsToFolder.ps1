<#
  .SYNOPSIS
    Backup Report Server (SSRS/PBIRS) items to a folder
  .DESCRIPTION
    This script will create a folder for each Report Server folder and save the contents of each folder to a subfolder of the output directory. 
    Currently must be run in Windows PowerShell (version 5.1) because the ReportingServicesTools module is not yet compatible with PowerShell Core (version 6+)
  .INPUTS
    - Parameters are currently the only way to pass input to this script.
    - Pipeline inputs are not yet supported.
  .OUTPUTS
    - Backup files are saved to the location specified by the OutputDirectory parameter.
    - Console output is silent by default. Verbose output can be displayed during the backup process by passing the -Verbose switch.
    - Pipeline outputs are not yet supported.
  .PARAMETER ServerName (Required)
    The DNS name or IP address of the Report Server (e.g., "MyReportServer", "MyReportServer.MyDomain.com", "127.0.0.1", etc.)
  .PARAMETER ServerPort (Optional)
    The port number of the Report Server (Defaults to 443)
  .PARAMETER RsInstance (Optional)
    The Report Server instance name (Defaults to "ReportServer")
  .PARAMETER RsRoot (Optional)
    The root folder of the Report Server (Defaults to "/")
  .PARAMETER OutputDirectory (Optional)
    The directory to save the backup to (Defaults to $env:TEMP\{ServerNameParameterValue}_{RsInstanceParameterValue}_Backup_{CurrentDate}_{CurrentTime})
    So, if you run this script with the default OutputDirectory and RsInstance parameter values on a server named "MyReportServer" on 2023-11-08 at 10:30:00, 
    the output directory will be: C:\Users\{YourUserName}\AppData\Local\Temp\MyReportServer_ReportServer_Backup_20231108_103000
  .PARAMETER OpenOutput (Optional)
    This is a switch parameter. When specified, the output directory will be opened after the script completes.
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer"
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer" -ServerPort 8080
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer" -OutputDirectory "C:\Temp\ReportServer_Backup"
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer" -RsRoot "/MyCustomRootFolder"
  .EXAMPLE
    Backup-ReportServerItemsToFolder.ps1 -ServerName "MyReportServer" -RsInstance "MyCustomInstance" -OpenOutput
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
    Version:  1.1
    Author:   James D. Bartlett III (@jamesdbartlett3@techhub.social)
    Date:     2024-10-22
    Acknowledgements:
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>

#Requires -Module ReportingServicesTools
#Requires -Version 5.1

param(
  [Parameter(Mandatory)][string]$ServerName
  , [Parameter()][int]$ServerPort = 443
  , [Parameter()][string]$RsInstance = "ReportServer"
  , [Parameter()][string]$RsRoot = "/"
  , [Parameter()][string]$OutputDirectory = (Join-Path -Path $env:TEMP -ChildPath "$($ServerName)_$($RsInstance)_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')")
  , [Parameter()][switch]$OpenOutput
)

### Note: This is currently commented out because the ReportingServicesTools module is not yet compatible with PowerShell Core.
### When the module has been updated to be compatible with PowerShell Core, you can uncomment this block and remove the `Requires -Version 5.1` line above the `param` block.
## If PowerShell version is greater than 5, import ReportingServicesTools module with -UseWindowsPowerShell parameter
# if ($PSVersionTable.PSVersion.Major -gt 5) {
#   Remove-Module ReportingServicesTools | Out-Null
#   try {
#     Import-Module ReportingServicesTools -UseWindowsPowerShell
#   } catch {
#     Write-Error "Unable to import ReportingServicesTools module with -UseWindowsPowerShell parameter. Please install ReportingServicesTools module in Windows PowerShell and try again."
#     Write-Host "Try: powershell.exe -NoProfile -ExecutionPolicy Bypass -Command 'Install-Module ReportingServicesTools -Scope CurrentUser'"
#     exit
#   }
# }

# Declare Report Server URI
$sourceRsUri = "https://$($ServerName):$($ServerPort)/$($RsInstance)/"

# Declare Proxy
$proxy = New-RsWebServiceProxy -ReportServerUri $sourceRsUri

# Get all catalog items NOT in subfolders of "/Users Folders"
$proxy.ListChildren("/", $false) | Where-Object { $_.TypeName -eq "Folder" -and $_.Path -notlike "/Users Folders*" } | ForEach-Object {
  Write-Verbose "Processing folder $($_.Path)..."
  $subfolder = (Join-Path -Path $OutputDirectory -ChildPath $_.Path)
  if (!(Test-Path -Path $subfolder)) {
    New-Item -Path $subfolder -ItemType Directory | Out-Null
  }
  Out-RsFolderContent -Proxy $proxy -RsFolder $_.Path -Recurse -Destination $subfolder
}

# Open output directory
if ($OpenOutput) {
  Invoke-Item -Path $OutputDirectory
}