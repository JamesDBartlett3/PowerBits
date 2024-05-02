#Requires -PSEdition Core
Function Copy-PowerBIReportContentToBlankPBIXFile {
  <#
  .SYNOPSIS
    Copies the contents of a published Power BI report into a new report published from a blank PBIX file.
  .DESCRIPTION
    This script will copy the contents of a published Power BI report into a new report published from a blank PBIX. This solves the problem where a Power BI report originally created in the web browser cannot be downloaded from the Power BI service as a PBIX file.
  .PARAMETER SourceReportId
    The ID of the report to copy from
  .PARAMETER SourceWorkspaceId
    The ID of the workspace to copy from
  .PARAMETER TargetReportId
    The ID of the report to copy to
  .PARAMETER TargetWorkspaceId
    The ID of the workspace to copy to
  .PARAMETER BlankPbix 
    Local path (or URL) to a blank PBIX file to upload and copy the source report's contents into
  .PARAMETER OutFile 
    Local path to save the new PBIX file to
  .EXAMPLE
    .\Copy-PowerBIReportContentToBlankPBIXFile.ps1 -SourceReportId "12345678-1234-1234-1234-123456789012" -SourceWorkspaceId "12345678-1234-1234-1234-123456789012" -TargetReportId "12345678-1234-1234-1234-123456789012" -TargetWorkspaceId "12345678-1234-1234-1234-123456789012"
  .NOTES
    This script does NOT require Azure AD app registration, service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must be allowed to download report PBIX files (see: "Download reports" setting in the Power BI Admin Portal).
      - The user must have "Contributor" or higher permissions on the source and target workspace(s).
    TODO
      - [ValidateScript({Test-Path $_})][string]$path on all file paths
      - Testing
      - Add usage, help, and examples.
      - Rename the function to something more accurate to its current capabilities.
      - [gc]::Collect() to free up memory
    ACKNOWLEDGEMENTS
      - This PS script was inspired by a blog article written by 
        one of the top minds in the Power BI space, Mathias Thierbach.
        Check out his article here: https://bit.ly/37ofVou
        And if you're not already using his pbi-tools for Power BI
        version control, you should check it out: https://pbi.tools
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits/blob/main/PowerShell/Scripts/Copy-PowerBIReportContentToBlankPBIXFile.ps1)
  .LINK
    [The author's blog](https://datavolume.xyz)
  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)
  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)
  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)
#>
  #Requires -Modules MicrosoftPowerBIMgmt
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true)][string]$SourceReportId,
    [Parameter(Mandatory = $true)][string]$SourceWorkspaceId,
    [Parameter(Mandatory = $false)][string]$TargetReportId,
    [Parameter(Mandatory = $false)][string]$TargetWorkspaceId = $SourceWorkspaceId,
    [Parameter(Mandatory = $false)][string]$BlankPbix,
    [Parameter(Mandatory = $false)][string]$OutFile
  )
  $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::New()
  [string]$blankPbixTempFile = Join-Path -Path $env:TEMP -ChildPath 'blank.pbix'
  [array]$validPbixContents = @('Layout', 'Metadata')
  [bool]$blankPbixIsUrl = $BlankPbix.StartsWith('http')
  [bool]$localFileExists = Test-Path $BlankPbix
  [bool]$remoteFileIsValid = $false
  [bool]$localFileIsValid = $false
  [bool]$defaultFileIsValid = $false
  Function FileIsBlankPbix($file) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
    $fileIsPbix = @($validPbixContents | Where-Object { $zip.Entries.Name -Contains $_ }).Count -gt 0
    $fileIsBlank = (Get-Item $file).length / 1KB -lt 20
    $zip.Dispose()
    if ($fileIsPbix -and $fileIsBlank) {
      Write-Debug "$file is a valid blank pbix file."
      return $true
    }
    else {
      Write-Error "$file is NOT a valid blank pbix file."
      return $false
    }
  }
  # If user did not specify a target report ID, use a blank PBIX file
  if (!$TargetReportId) {
    # If user specified a URL to a file, download and validate it as a blank PBIX file
    if ($blankPbixIsUrl) {
      Write-Debug "Downloading file: $BlankPbix..."
      Invoke-WebRequest -Uri $BlankPbix -OutFile $blankPbixTempFile
      Write-Debug 'Validating downloaded file...'
      $remoteFileIsValid = FileIsBlankPbix($blankPbixTempFile)
    }
    # If user specified a local path to a file, validate it as a blank PBIX file
    elseif ($localFileExists) {
      Write-Debug "Validating user-supplied file: $BlankPbix..."
      $localFileIsValid = FileIsBlankPbix($BlankPbix)
    }
    # If user didn't specify a blank PBIX file, check for a valid blank PBIX in the temp location
    elseif (Test-Path $blankPbixTempFile) {
      Write-Debug "Validating pbix file found in temp location: $blankPbixTempFile..."
      $defaultFileIsValid = FileIsBlankPbix($blankPbixTempFile)
    }
    # If user did not specify a blank PBIX file, and a valid blank PBIX is not in the temp location,
    # download one from GitHub and check if it's valid and blank
    else {
      Write-Debug "Downloading a blank pbix file from GitHub to $blankPbixTempFile..."
      $BlankPbixUri = 'https://github.com/JamesDBartlett3/PowerBits/raw/main/Misc/blank.pbix'
      Invoke-WebRequest -Uri $BlankPbixUri -OutFile $blankPbixTempFile
      $defaultFileIsValid = FileIsBlankPbix($blankPbixTempFile)
    }
    # If we downloaded a valid blank PBIX file, use it.
    if ($remoteFileIsValid -or $defaultFileIsValid) {
      $BlankPbix = $blankPbixTempFile
    }
    # If a valid blank PBIX file could not be obtained by any of the above methods, throw an error.
    if (!$TargetReportId -and !$localFileIsValid -and !$remoteFileIsValid -and !$defaultFileIsValid) {
      Write-Error 'No targetReportId specified & no valid blank PBIX file found. Please specify one or the other.'
      return
    }
    [bool]$pbixIsValid = ($localFileIsValid -or $remoteFileIsValid -or $defaultFileIsValid)
  }
  try {
    $headers = Get-PowerBIAccessToken
  }
  catch {
    Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID authentication dialog...' -ForegroundColor DarkYellow
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
  }
  finally {
    Write-Host '🔑 Power BI Access Token acquired.' -ForegroundColor Green
    Write-Debug "Target Report ID is null: $(!$TargetReportId)"
    $pbiApiBaseUri = 'https://api.powerbi.com/v1.0/myorg'
    # If a valid blank PBIX was found, publish it to the target workspace
    if ($pbixIsValid) {
      Write-Debug "Publishing $BlankPbix to target workspace..."
      $publishResponse = New-PowerBIReport -Path $BlankPbix -WorkspaceId $TargetWorkspaceId -ConflictAction CreateOrOverwrite
      Write-Debug "Response: $publishResponse"
      $TargetReportId = $publishResponse.Id
    }
    # Assemble the UpdateReportContent API URI and request body
    $updateReportContentEndpoint = "$pbiApiBaseUri/groups/$TargetWorkspaceId/reports/$TargetReportId/UpdateReportContent"
    $body = @"
    {
      "sourceReport": {
        "sourceReportId": "$SourceReportId",
        "sourceWorkspaceId": "$SourceWorkspaceId"
      },
      "sourceType": "ExistingReport"
    }
"@
    # Update the target report with the source report's content
    $headers.Add('Content-Type', 'application/json')
    $response = Invoke-RestMethod -Uri $updateReportContentEndpoint -Method POST -Headers $headers -Body $body
    # If user did not specify an output file, use the source report's name
    $sourceReportName = (Get-PowerBIReport -Id $SourceReportId -WorkspaceId $SourceWorkspaceId).Name
    $OutFile = !!$OutFile ? $OutFile : "$($sourceReportName)_Clone.pbix"
    # Export the target report to a PBIX file
    Export-PowerBIReport -WorkspaceId $TargetWorkspaceId -Id $response.id -OutFile $OutFile
    # Assemble the Datasets API URI
    $datasetsEndpoint = "$pbiApiBaseUri/groups/$TargetWorkspaceId/datasets"
    # Delete the target dataset and report from the target workspace
    Invoke-RestMethod "$datasetsEndpoint/$($response.datasetId)" -Method DELETE -Headers $headers
  }
}
Function Export-PowerBIReportsFromWorkspaces {
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
        Write-Host '🔒 Power BI Access Token required. Launching Microsoft Entra ID authentication dialog...' -ForegroundColor DarkYellow
        Start-Sleep -s 1
        Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
        $headers = Get-PowerBIAccessToken
      }
      if (!$headers) {
        Write-Host '❌ Power BI Access Token not acquired. Exiting...' -ForegroundColor Red
        Exit
      }
    }
    Write-Host '🔑 Power BI Access Token acquired.' -ForegroundColor Green
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
}
