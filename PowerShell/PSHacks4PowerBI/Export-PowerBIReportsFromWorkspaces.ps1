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
    - Refactor target directory selection to use terminal prompt
    - Add "extractWithPbiTools" boolean parameter
      - Implement "extractWithPbiTools" parameter
    - Add option to overwrite existing report files
    - Add rdl support
    - Add usage, help, and examples
    - Change "error_log_(timestamp).txt" -- keep logs from previous runs
    - Remove all "testing" code

  .ACKNOWLEDGEMENTS
    -

#>

Function Export-PowerBIReportsFromWorkspaces {

  #Requires -PSEdition Core
  #Requires -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

  [CmdletBinding()]
  Param(
    [parameter(Mandatory = $true)][string]$destinationFolder,
    [parameter(Mandatory = $false)][int]$throttleLimit = 1,
    [parameter(Mandatory = $false)][switch]$extractWithPbiTools
  )

  [int]$waitSeconds = 30
  $currentDateTime = Get-Date -UFormat "%Y%m%d_%H%M%S"
  
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
    Out-Debug "Headers: `n" $headers
    [array]$ignoreWorkspaces = @(
      "COVID-19 Tracking Report"
      , "COVID-19 US Tracking Report"
      , "Gen2 Utilization Metrics"
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

    # If user didn't specify a destination folder, use the standard temp directory
    $targetDir = !$destinationFolder ? $env:TEMP : $destinationFolder

    # Create a log file to record errors
    $errorLog = "$targetDir\error_log_$currentDateTime.txt"

    # TODO: Handle report type (rdl vs pbix)
      # $_.WebUrl -contains "/rdlreports/"
      # $_.WebUrl -contains "/reports/"
    ForEach ($w in $workspaces) {
      $workspaceID = $w.Id
      $workspaceName = $w.Name
      $reports = Get-PowerBIReport -WorkspaceId $workspaceID |
      Where-Object {
        $_.Name -notIn $ignoreReports
      } |
      Sort-Object -Property Name

      if ($reports -like "*Unauthorized*") {
        Add-Content -LiteralPath $errorLog "Error on $workspaceName workspace: Unauthorized."
      }

      if (-not (Test-Path -LiteralPath "$targetDir\$workspaceName" -PathType Container)) {
        New-Item -LiteralPath "$targetDir\$workspaceName" -ItemType Directory | Out-Null
      }

      $reports | ForEach-Object -Parallel {
        $waitSeconds = $using:waitSeconds
        $throttleLimit = $using:throttleLimit
        $reportID = $_.Id
        $reportName = $_.Name
        $errorLog = $using:errorLog
        $targetDir = $using:targetDir
        $workspaceID = $using:workspaceID
        $workspaceName = $using:workspaceName
        $targetReportDir = "$targetDir\$workspaceName\$reportName"
        $targetFile = "$targetReportDir\$reportName.pbix"

        if (-not (Test-Path -LiteralPath $targetReportDir -PathType Container)) {
          New-Item -LiteralPath $targetReportDir -ItemType Directory | Out-Null
        }
        Set-Location -LiteralPath $targetReportDir
        Write-Verbose "_______________________________________________________"
        Write-Verbose "Exporting $reportName to $targetFile..."

        if (Test-Path -LiteralPath $targetFile) {
          Write-Verbose "$targetFile already exists; Skipping."
        }
        else {
          $message = Export-PowerBIReport -WorkspaceId $workspaceID -Id $reportID -OutFile ".\$reportName.pbix" 2>&1 |
          Out-String

          if ($message -like "*BadRequest*") {
            $message = "Incremental Refresh"
          }
          elseif ($message -like "*NotFound*" -or $message -like "*Forbidden*") {
            $message = "Downloads Disabled"
          }
          elseif ($message -like "*TooManyRequests*") {
            $message = "Reached Power BI API Rate Limit; Try Again Later"
          }
          elseif ($message -like "*Unauthorized*") {
            $message = "Unauthorized"
          }
          else { $message = "Done" }

          $fullPathMessage = "$targetFile" + ": $message"
          $shortPathMessage = "$workspaceName\$reportName" + ": $message"

          if ($message -ne "Done") {
            Add-Content -LiteralPath $errorLog $fullPathMessage
            Write-Output "❌ `e[38;2;255;0;0m$shortPathMessage (see $errorLog for details)`e[0m" # Red
          }
          else {
            Write-Output "✔ $shortPathMessage"
          }
          Write-Verbose "_______________________________________________________"
          # Write-Verbose "Waiting $waitSeconds seconds to avoid hitting the Power BI API Rate Limit (200 req/hr)..."
          Start-Sleep $waitSeconds
        }
        Set-Location -LiteralPath $targetDir
      } -ThrottleLimit $throttleLimit
    }

    # Remove any empty directories
    Get-ChildItem $targetDir -Recurse -Attributes Directory |
    Where-Object { $_.GetFileSystemInfos().Count -eq 0 } |
    Remove-Item

    Invoke-Item $targetDir

  }

}