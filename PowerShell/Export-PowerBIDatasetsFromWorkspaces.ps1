<#

  .SYNOPSIS
    Function: Export-PowerBIDatasetsFromWorkspaces
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

  .DESCRIPTION
    Export Power BI datasets from multiple workspaces in parallel

  .EXAMPLE
    Export-PowerBIDatasetsFromWorkspaces -OutputFolder C:\Datasets -ExtractWithPbiTools -SkipExistingFiles -ThrottleLimit 10

  .PARAMETER OutputFolder
    The folder where the datasets will be saved. If the folder does
    not exist, it will be created.

  .PARAMETER ExtractWithPbiTools
    If specified, exported PBIX datasets will be extracted with 
    pbi-tools after they are exported. Requires pbi-tools to be
    installed. See: https://pbi.tools

  .PARAMETER SkipExistingFiles
    If specified, existing files will be skipped. If not specified,
    existing files will be overwritten.

  .PARAMETER ThrottleLimit
    The maximum number of datasets that will be exported in parallel.
    Defaults to 1.

  .NOTES
    This function does NOT require Azure AD app registration, 
    service principal creation, or any other special setup.
    The only requirements are:
      - The user must be able to run PowerShell (and install the
        MicrosoftPowerBIMgmt module, if it's not already installed).
      - The user must be allowed to download report PBIX files
        (see: "Download reports" setting in the Power BI Admin Portal).
    
    TODO


#>

Function Export-PowerBIDatasetsFromWorkspaces {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory=$false)][string]$OutputFolder,
    [Parameter(Mandatory=$false)][switch]$ExtractWithPbiTools,
    [Parameter(Mandatory=$false)][switch]$SkipExistingFiles,
    [Parameter(Mandatory=$false)][int]$ThrottleLimit = 1
  )

  #Requires -PSEdition Core -Modules MicrosoftPowerBIMgmt, Microsoft.PowerShell.ConsoleGuiTools

  [string]$currentDateTime = Get-Date -UFormat "%Y%m%d_%H%M%S"
  [string]$fallbackDir = Join-Path -Path $env:TEMP -ChildPath "PowerBIWorkspaces"
  
  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

  Function Convert-PbixToProj {
    Param(
      [Parameter(Mandatory = $true)][string]$PbixPath,
      [Parameter(Mandatory = $true)][string]$ShortPath
    )
    try {
      Invoke-Expression pbi-tools | Out-Null
    } catch {
      Write-Error "'pbi-tools' command not found. See: https://pbi.tools"
      Write-Warning $Error[0]
    }
    finally{
      if (!$Error[0]) {
        $command = "pbi-tools extract -pbixPath ""$PbixPath"""
        Write-Debug "Running command: $command"
        Write-Output "📦 Extracting: $ShortPath"
        Invoke-Expression $command | Out-Null
      }
    }
  }

  $fn_PbixToProj = ${function:Convert-PbixToProj}.ToString()

  try {
    $headers = Get-PowerBIAccessToken
  } catch {
    Write-Output "Power BI Access Token required. Launching authentication dialog..."
    Start-Sleep -s 1
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue | Out-Null
    $headers = Get-PowerBIAccessToken
  } finally {
  
  }

}