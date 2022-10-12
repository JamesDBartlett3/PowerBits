<# 

.SYNOPSIS
    Title: Import-PSHacks4PowerBIFunctions
    Author: @JamesDBartlett3 (James D. Bartlett III)

.DESCRIPTION
    - Imports functions for the "PowerShell Hacks for Power BI" demo session

.PARAMETERS
    - 

.TODO
    - Add all needed functions
    - Parameterize with list of functions to import?

.EXAMPLE
    Import-PSHacks4PowerBIFunctions.ps1

#>

# List of functions to import
$functions = @(
    "Get-UserDatasets.ps1"
    , "Match-DatasetsWithWorkspaces.ps1"
    , "Takeover-UserDataset.ps1"
    , "Copy-PowerBIReportContentToBlankPBIXFile.ps1"
    , "Get-AADGroupMembers.ps1"
    # , "Export-ReportsFromWorkspaces-Parallel.ps1"
    # , "Audit-PowerBIWorkspaceSecurity.ps1"
    # , "Push-DataToPowerBIStreamingDataset_Demo.ps1"
    # , "Add-EntityToWorkspaces.ps1"
    # , "Remove-EntityFromWorkspaces.ps1"
)

# Dotsource all functions in current directory whose names match those in $functions array
Get-ChildItem -Path $PSScriptRoot -Filter *.ps1 -Recurse |
    Where-Object {$_.Name -in $functions -and $_.FullName -ne $PSCommandPath} |
    ForEach-Object {
        Write-Output "Importing `e[38;2;0;255;0m$($_.BaseName)`e[0m..."
        . $_.FullName
    }