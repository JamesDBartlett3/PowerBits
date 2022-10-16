<#

  .SYNOPSIS
    Function: Export-ReportsFromWorkspaces-Parallel
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
      - The user must have permissions to access the workspace(s)
        in the Power BI service.

  .TODO
    - Convert to function
    - Replace Out-Gridview with ConsoleGuiTools version
    - Add option to overwrite existing report files
    - Add usage, help, and examples
    - Change "error_log_(timestamp).txt" -- keep logs from previous runs
    - Re-implement testing logic

  .ACKNOWLEDGEMENTS
    -

#>

#Requires -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

[bool]$testing = $False
[int]$waitSeconds = 30
[int]$throttleLimit = 10


$token = $null
$token = Get-PowerBIAccessToken
if (!$token){
  Connect-PowerBIServiceAccount | Out-Null
}


$ignoreWorkspaces = @(
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
$ignoreReports = @(
  "Report Usage Metrics Report"
  , "Dashboard Usage Metrics Report"
)

if ($testing) {
  $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
    Where-Object {
                  $_.Type -eq "Workspace" -and
                  $_.Name -eq "Testing" -and
                  $_.State -eq "Active"
                }
  $targetDir = "C:\temp"
} else {
  Function Get-TargetDirectory($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.SelectedPath = $initialDirectory
    $FolderBrowserDialog.Description = "Select Target Directory"
    $FolderBrowserDialog.ShowDialog() | Out-Null
    $FolderBrowserDialog.SelectedPath
  }
  $workspaces = Get-PowerBIWorkspace -Scope Organization -All |
    Where-Object {
                  $_.Type -eq "Workspace" -and
                  $_.State -eq "Active" -and
                  $_.Name -notIn $ignoreWorkspaces
                } |
    Select-Object Id, Name |
    Sort-Object -Property Name |
    Out-GridView -PassThru -Title "Select Workspaces (Ctrl+Click to Multi-Select)"
  $targetDir = Get-TargetDirectory(Get-Location)
}

$errorLog = "$targetDir\error_log.txt"
Clear-Content -LiteralPath $errorLog

ForEach($w in $workspaces) {
  $workspaceID = $w.Id
  $workspaceName = $w.Name
  $reports = Get-PowerBIReport -WorkspaceId $workspaceID |
    Where-Object {
                  $_.Name -notIn $ignoreReports -and
                  $_.WebUrl -notLike "*rdlreports*"
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
    } else {
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
      else {$message = "Done"}

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

if (-not $testing) {
  Disconnect-PowerBIServiceAccount | Out-Null
}

# Remove any empty directories
Get-ChildItem $targetDir -Recurse -Attributes Directory |
  Where-Object {$_.GetFileSystemInfos().Count -eq 0} |
  Remove-Item

Invoke-Item $targetDir
