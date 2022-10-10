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
    "Get-UserDatasets.ps1",
    "Takeover-UserDataset.ps1",
    "Match-DatasetsWithWorkspaces.ps1",
    "Export-ReportsFromWorkspaces-Parallel.ps1",
    "Copy-PowerBIReportContentToBlankPBIXFile.ps1",
    "Audit-PowerBIWorkspaceSecurity.ps1",
    "Push-DataToPowerBIStreamingDataset_Demo.ps1",
    "Add-EntityToWorkspaces.ps1",
    "Remove-EntityFromWorkspaces.ps1"
)

# Get list of functions in current directory whose names match those in $functions array
$functionFiles = Get-ChildItem -Path . -Filter *.ps1 -Recurse |
    Where-Object -Property Name -In $functions

# Dotsource all functions in $functionFiles
ForEach($f in $functionFiles) {
    . $f.FullName
}