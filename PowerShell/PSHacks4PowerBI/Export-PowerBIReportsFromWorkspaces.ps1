<#

  .SYNOPSIS
    Function: Export-PowerBIReportsFromWorkspaces
    Author: @JamesDBartlett3 (James D. Bartlett III)

  .DESCRIPTION
    Export Power BI reports from multiple workspaces in parallel

  .PARAMETERS
    - 

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must be allowed to download report PBIX files
        (see: "Download reports" setting in the Power BI Admin Portal).

  .TODO
    - Add "extractWithPbiTools" boolean parameter
      - Implement "extractWithPbiTools" parameter
    - Replace $waitSeconds with a more robust wait mechanism
      - Use pbimonitor scripts for inspiration
    - Add usage, help, and examples

  .ACKNOWLEDGEMENTS
    -

#>

Function Export-PowerBIReportsFromWorkspaces {

  #Requires -PSEdition Core
  #Requires -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

  [CmdletBinding()]
  Param(
    [parameter(Mandatory = $false)][string]$destinationFolder = $null,
    [parameter(Mandatory = $false)][int]$throttleLimit = 1,
    [parameter(Mandatory = $false)][switch]$extractWithPbiTools,
    [parameter(Mandatory = $false)][switch]$skipExistingFiles
  )

  [string]$currentDateTime = Get-Date -UFormat "%Y%m%d_%H%M%S"
  [string]$fallbackDir = Join-Path -Path $env:TEMP -ChildPath "PowerBIWorkspaces"
  
  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

  try {
    $headers = Get-PowerBIAccessToken
  }

  catch {
    Write-Output "Power BI Access Token required. Launching authentication dialog..."
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
  }

  finally {
    Write-Debug "Headers: `n $($headers.Keys)`n $($headers.Values)"
    [array]$ignoreWorkspaces = @(
      "Gen2 Utilization Metrics"
      , "Azure DevOps Dashboard"
      , "Microsoft Project Web App"
      , "Office365 Usage Analytics"
      , "Power BI Premium Capacity Metrics"
      , "Microsoft 365 Usage Analytics"
      , "Dataflow Snapshots"
    )
    [array]$ignoreReports = @(
      "Report Usage Metrics Report"
      , "Dashboard Usage Metrics Report"
    )

    # Get list of workspaces and prompt user to select which ones to export
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
      Where-Object {
        $_.Type -eq "Workspace" -and
        $_.State -eq "Active" -and
        $_.Name -notIn $ignoreWorkspaces
      } |
      Select-Object Name, Id |
      Sort-Object -Property Name |
      Out-ConsoleGridView -Title "Select Workspaces to Export"

    # If user didn't specify a destination folder, fall back to $fallbackDir
    $targetDir = if(!$destinationFolder) { $fallbackDir } else { $destinationFolder }
    Write-Output "Target directory: $targetDir"

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
      if ($reports -like "*Unauthorized*") { #TODO: Proper error handling
        Add-Content -LiteralPath $errorLog "Error on $workspaceName workspace: Unauthorized."
      }

      # Declare $workspacePath variable and create workspace folder if it doesn't exist
      $workspacePath = Join-Path -Path $targetDir -ChildPath $workspaceName
      if (!(Test-Path -LiteralPath $workspacePath -PathType Container)) {
        New-Item -Path $workspacePath -ItemType Directory | Out-Null
      }

      # Loop through all reports in the current workspace and download them
      $reports | ForEach-Object -Parallel {
        $DebugPreference = $using:DebugPreference # bug work around
        $VerbosePreference = $using:VerbosePreference # bug work around
        $InformationPreference = $using:InformationPreference # bug work around
        $reportID = $_.Id
        $reportName = $_.Name
        $reportWebUrl = $_.WebUrl
        $errorLog = $using:errorLog
        $targetDir = $using:targetDir
        $workspaceID = $using:workspaceID
        $workspaceName = $using:workspaceName
        $workspacePath = $using:workspacePath
        $targetReportPathBaseName = Join-Path -Path $workspacePath -ChildPath $reportName
        $shortPathBaseName = Join-Path -Path $workspaceName -ChildPath $reportName
        $targetFilePath, $shortPath = ($reportWebUrl -like "*/rdlreports/*") ?
          "$targetReportPathBaseName.rdl", "$shortPathBaseName.rdl" :
          "$targetReportPathBaseName.pbix", "$shortPathBaseName.pbix"
        Write-Debug "Report WebUrl: $reportWebUrl"
        Write-Verbose "_______________________________________________________"
        Write-Verbose "Exporting $reportName to $targetFilePath..."

        # If user specified to skip existing files, check if the file exists
        if ((Test-Path $targetFilePath) -and $using:skipExistingFiles) {
          Write-Output "⤵️  $shortPath already exists; Skipping..."
        }
        # Otherwise, download the report
        else {
          # If $targetFilePath already exists, remove it
          if (Test-Path $targetFilePath) { Remove-Item $targetFilePath -Force -ErrorAction SilentlyContinue }
          $message = Export-PowerBIReport -WorkspaceId $workspaceID -Id $reportID -OutFile $targetFilePath 2>&1 |
          Out-String

          # Error handling for Export-PowerBIReport 
          #TODO: Proper error handling
          $message = switch ($true) {
            { $message -like "*BadRequest*" } { "Incremental Refresh" }
            { $message -like "*NotFound*" -or $message -like "*Forbidden*" } { "Downloads Disabled" }
            { $message -like "*TooManyRequests*" } { "Reached Power BI API Rate Limit; Try Again Later." }
            { $message -like "*Unauthorized*" } { "Unauthorized" }
            { $true } { "Done" }
          }

          $fullPathMessage = "$targetFilePath`: $message"
          $shortPathMessage = "$shortPath`: $message"

          if ($message -ne "Done") {
            Add-Content -LiteralPath $errorLog $fullPathMessage
            Write-Output "❌ `e[38;2;255;0;0m$shortPathMessage (see $errorLog for details)`e[0m" # Red
          } 
          else { Write-Output "✅ $shortPathMessage" }
          Write-Verbose "_______________________________________________________"

        }
        Set-Location -LiteralPath $targetDir
      } -ThrottleLimit $throttleLimit
    }

    # Remove any empty directories
    Get-ChildItem $targetDir -Recurse -Attributes Directory |
      Where-Object { $_.GetFileSystemInfos().Count -eq 0 } | Remove-Item

  }

}