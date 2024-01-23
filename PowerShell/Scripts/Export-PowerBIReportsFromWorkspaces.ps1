<#
  .SYNOPSIS
    Exports Power BI reports (.pbix and .rdl) from Power BI workspaces to a local folder.
  
  .DESCRIPTION
    This script will export Power BI reports (.pbix and .rdl) from Power BI workspaces to a local folder.
    Optional features:
    - Extract the source code of exported PBIX files using pbi-tools.
    - Skip existing files to avoid overwriting them.
    - Export one report at a time or in parallel (default behavior: count processor cores and run that many parallel processes).
  
  .PARAMETER OutputFolder
    The folder where the reports will be saved. If the folder does not exist, it will be created.
  
  .PARAMETER ExtractWithPbiTools
    If specified, exported PBIX reports will be extracted with pbi-tools after they are exported. Requires pbi-tools to be installed. See: https://pbi.tools
  
  .PARAMETER SkipExistingFiles
    If specified, existing files will be skipped. If not specified, existing files will be overwritten.
  
  .PARAMETER ThrottleLimit
    The maximum number of reports that will be exported in parallel. Defaults to the number of processor cores detected.
  
  .EXAMPLE
    # Export reports to the default folder in the temp directory, overwriting any existing files there
    .\Export-PowerBIReportsFromWorkspaces.ps1
  
  .EXAMPLE
    # Export reports, up to two at a time, to the "C:\Reports" folder, skip any files that already exist there, 
    # and use pbi-tools to extract the source code of the PBIX files into subfolders named after the reports they came from
    .\Export-PowerBIReportsFromWorkspaces.ps1 -OutputFolder C:\Reports -ExtractWithPbiTools -SkipExistingFiles -ThrottleLimit 2
  
  .NOTES
    This script does NOT require Azure AD app registration, service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must be allowed to download report PBIX files (see: "Download reports" setting in the Power BI Admin Portal).
    
    TODO
      - [ValidateScript({Test-Path $_})][string]$path on all file paths
      - Add ability to find and export report-less datasets
      - Fix bug where reports with illegal characters in name cannot be extracted
      - Add $workspacesToExport parameter to allow user to specify
        which workspaces to export from.
        - This would require a change to the Get-PowerBIWorkspace
          function to allow filtering by workspace name.
      - Add dynamic rate limiting to avoid throttling
        - Use pbimonitor scripts for inspiration
        - https://github.com/RuiRomano/pbiscripts/blob/main/Workspace-TenantScan.ps1
      - Add logic to spread parallelism over multiple workspaces
      - Experiment with using classes (https://bit.ly/3glYGZf)
        to improve parallelism performance
      - Add usage, help, and examples
      - [gc]::Collect() to free up memory
    
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
  
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits/blob/main/PowerShell/Scripts/Export-PowerBIReportsFromWorkspaces.ps1)
  
  .LINK
    [The author's blog](https://datavolume.xyz)
    
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
#>

#Requires -PSEdition Core
#Requires -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

[CmdletBinding()]
Param(
  [parameter(Mandatory = $false)][string]$OutputFolder,
  [parameter(Mandatory = $false)][switch]$ExtractWithPbiTools,
  [parameter(Mandatory = $false)][switch]$SkipExistingFiles,
  [parameter(Mandatory = $false)][int]$ThrottleLimit = [Environment]::ProcessorCount
)

begin {
  # Declare the servicePrincipal global variables
  $global:servicePrincipalId = $null
  $global:servicePrincipalTenantId = $null
  $global:servicePrincipalSecret = $null
  $global:credential = $servicePrincipalId ? (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $servicePrincipalId, ($servicePrincipalSecret | ConvertTo-SecureString -AsPlainText -Force)) : $null
  [string]$currentDateTime = Get-Date -UFormat "%Y%m%d_%H%M%S"
  [string]$fallbackDir = Join-Path -Path $env:TEMP -ChildPath "PowerBIWorkspaces"
  $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()

  Function Convert-PbixToProj {
    Param(
      [Parameter(Mandatory = $true)][string]$PbixPath,
      [Parameter(Mandatory = $true)][string]$ShortPath
    )
    try {
      Invoke-Expression pbi-tools | Out-Null
    }
    catch {
      Write-Error "'pbi-tools' command not found. See: https://pbi.tools/tutorials/getting-started-cli.html"
      Write-Warning $Error[0]
    }
    finally {
      if (!$Error[0]) {
        $command = "pbi-tools extract -pbixPath ""$PbixPath"""
        Write-Debug "Running command: $command"
        Write-Host "📦 Extracting: $ShortPath"
        Invoke-Expression $command | Out-Null
      }
    }
  }
  $fn_PbixToProj = ${function:Convert-PbixToProj}.ToString()
}

process{

  try {
    $headers = Get-PowerBIAccessToken
  }
  catch {
    if ($servicePrincipalId) {
      Connect-PowerBIServiceAccount -ServicePrincipal -Tenant $servicePrincipalTenantId -Credential $credential
      $headers = Get-PowerBIAccessToken
    }
    else {
      Write-Host '🔒 Power BI Access Token required. Launching Azure Active Directory authentication dialog...'
      Start-Sleep -s 1
      Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
      $headers = Get-PowerBIAccessToken
    }
    if ($headers) {
      Write-Host '🔑 Power BI Access Token acquired. Proceeding...'
    }
    else {
      Write-Host '❌ Power BI Access Token not acquired. Exiting...'
      exit
    }
  }
  
  Write-Host '🔑 Power BI Access Token acquired.'
  
  # If debugging, display the access token
  Write-Debug "Headers: `n $($headers.Keys)`n $($headers.Values)"
  
  # Get names of Workspaces and Reports to ignore from IgnoreList.json file
  # Most of these are template apps and/or auto-generated by Microsoft
  [PSCustomObject]$ignoreObjects = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "../IgnoreList.json") | ConvertFrom-Json
  [array]$ignoreWorkspaces = $ignoreObjects.IgnoreWorkspaces
  [array]$ignoreReports = $ignoreObjects.IgnoreReports
  
  # Get list of workspaces and prompt user to select which ones to export
  $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
  Where-Object {
    $_.Type -eq "Workspace" -and
    $_.State -eq "Active" -and
    $_.Name -notIn $ignoreWorkspaces
  } | Select-Object Name, Id | Sort-Object -Property Name |
  Out-ConsoleGridView -Title "Select Workspaces to Export"
  
  # If user didn't specify a destination folder, fall back to $fallbackDir
  $targetDir = $OutputFolder ? $OutputFolder : $fallbackDir
  Write-Host "Target directory: $targetDir"
  
  # If target directory doesn't exist, create it
  if (!(Test-Path -LiteralPath $targetDir)) {
    New-Item -Path $targetDir -ItemType Directory | Out-Null
  }
  
  # Create a log file to record errors
  $errorLog = Join-Path -Path $targetDir -ChildPath "error_log_$currentDateTime.txt"
  
  # Open $targetDir in Windows Explorer
  Invoke-Item $targetDir
  
  # Loop through all selected workspaces and get list of reports in them
  ForEach ($w in $workspaces) {
    $workspaceID = $w.Id
    $workspaceName = $w.Name
    $reports = Get-PowerBIReport -WorkspaceId $workspaceID |
    Where-Object {
      $_.Name -notIn $ignoreReports
    } | Sort-Object -Property Name
    
    # If user does not have access to the current workspace, log an error and skip it
    #TODO: Proper error handling
    if ($reports -like "*Unauthorized*") {
      Add-Content -LiteralPath $errorLog "Error on $workspaceName workspace: Unauthorized."
    }
    
    # Declare $workspacePath variable and create workspace folder if it doesn't exist
    $workspacePath = Join-Path -Path $targetDir -ChildPath $workspaceName
    if (!(Test-Path -LiteralPath $workspacePath -PathType Container)) {
      New-Item -Path $workspacePath -ItemType Directory | Out-Null
    }
    
    # Loop through all reports in the current workspace and download them in parallel
    $reports | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    
      # Workaround for Write-Debug, Write-Verbose, and Write-Warning not working in parallel
      $DebugPreference = $using:DebugPreference 
      $VerbosePreference = $using:VerbosePreference 
      $InformationPreference = $using:InformationPreference
      
      # Declare variables for current report
      $reportID = $_.Id
      $reportName = $_.Name
      $reportWebUrl = $_.WebUrl
      $errorLog = $using:errorLog
      $targetDir = $using:targetDir
      $workspaceID = $using:workspaceID
      $workspaceName = $using:workspaceName
      $workspacePath = $using:workspacePath
      $SkipExistingFiles = $using:SkipExistingFiles
      ${function:Convert-PbixToProj} = $using:fn_PbixToProj
      $targetReportPathBaseName = Join-Path -Path $workspacePath -ChildPath $reportName
      $shortPathBaseName = Join-Path -Path $workspaceName -ChildPath $reportName
      $targetFilePath, $shortPath = ($reportWebUrl -like "*/rdlreports/*") ?
      "$targetReportPathBaseName.rdl", "$shortPathBaseName.rdl" :
      "$targetReportPathBaseName.pbix", "$shortPathBaseName.pbix"
      Write-Debug "Report WebUrl: $reportWebUrl"
      Write-Verbose "_______________________________________________________"
      Write-Verbose "Exporting $reportName to $targetFilePath..."
      
      # If user specified to skip existing files, check if the file exists
      if ((Test-Path -Path $targetFilePath) -and $SkipExistingFiles) {
        Write-Host "⤵️  $shortPath already exists; Skipping..."
      }
      # Otherwise, download the report
      else {
        # If $targetFilePath already exists, remove it
        if (Test-Path -Path $targetFilePath) { Remove-Item $targetFilePath -Force -ErrorAction SilentlyContinue }
        
        # Export the report and store the response in $message
        $message = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/groups/$workspaceID/reports/$reportID/Export" `
          -Method GET -Headers $using:headers `
          -ContentType "application/octet-stream" `
          -Body '{"preferClientRouting":true}' `
          -ErrorVariable message -ErrorAction SilentlyContinue `
          -OutFile $targetFilePath 2>&1 | Out-String
        
        # Error handling for Export-PowerBIReport 
        #TODO: proper error handling
        #TODO: rate limiting
        $message = switch ($true) {
          { $message -like "*BadRequest*" } { "Incremental Refresh" }
          { $message -like "*NotFound*" -or $message -like "*Forbidden*" -or $message -like "*Disabled*" } { "Downloads Disabled" }
          { $message -like "*TooManyRequests*" } { "Reached Power BI API Rate Limit; Try Again Later." }
          { $message -like "*Unauthorized*" } { "Unauthorized" }
          default { "Done" }
        }
        
        $fullPathMessage = "$targetFilePath`: $message"
        $shortPathMessage = "$shortPath`: $message"
        
        if ($message -ne "Done") {
          Add-Content -LiteralPath $errorLog $fullPathMessage
          Write-Host "❌ `e[38;2;255;0;0m$shortPathMessage (see $errorLog for details)`e[0m" # Red
        } 
        else { Write-Host "✅ $shortPathMessage" }
        Write-Verbose "_______________________________________________________"
      }
      if ($using:ExtractWithPbiTools -and $targetFilePath -like "*.pbix") {
        Convert-PbixToProj -PbixPath $targetFilePath -ShortPath $shortPath
      }
    }
    $headers = Get-PowerBIAccessToken
  }
  # Remove any empty directories
  Get-ChildItem $targetDir -Recurse -Attributes Directory |
  Where-Object { $_.GetFileSystemInfos().Count -eq 0 } | Remove-Item
}