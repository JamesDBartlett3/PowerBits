<# 
	.SYNOPSIS
		Title: Import-PSHacks4PowerBIFunctions
		Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

	.DESCRIPTION
		- Imports functions for the "PowerShell Hacks for Power BI" demo session

	.NOTES
		- Version: 1.0

	.EXAMPLE
		. .\Import-PSHacks4PowerBIFunctions.ps1
#>

# Pre-emptively import problematic modules
Import-Module Az.Resources -ErrorAction SilentlyContinue | Out-Null

# List of functions to import
$functions = @(
	"Get-UserDatasets.psm1"
	, "Join-DatasetsWithWorkspaces.psm1"
	, "Update-UserDatasetOwner.psm1"
	, "Copy-PowerBIReportContentToBlankPBIXFile.psm1"
	, "Get-DataGatewayNodesStatus.psm1"
	, "Export-PowerBIReportsFromWorkspaces.psm1"
  , "Export-PowerBIWorkspaceSecurity.psm1"
	, "Get-PowerBIBareDatasetsFromWorkspaces.psm1"
	, "Export-PowerBIBareDatasetFromWorkspace.psm1"
)

# Dotsource all functions in current directory whose names match those in $functions array
Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.psm1 |
	Where-Object { $_.Name -in $functions -and $_.FullName -ne $PSCommandPath } |
	ForEach-Object {
		Write-Output "Importing `e[38;2;0;255;0m$($_.BaseName)`e[0m..."
		Import-Module $_.FullName -Force
	}