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
	"Get-UserDatasets.ps1"
	, "Join-DatasetsWithWorkspaces.ps1"
	, "Update-UserDatasetOwner.ps1"
	, "Copy-PowerBIReportContentToBlankPBIXFile.ps1"
	, "Get-DataGatewayNodesStatus.ps1"
	, "Export-PowerBIReportsFromWorkspaces.ps1"
  , "Export-PowerBIWorkspaceSecurity.ps1"
	, "Get-PowerBIBareDatasetsFromWorkspaces.ps1"
	, "Export-PowerBIBareDatasetFromWorkspace.ps1"
)

# Dotsource all functions in current directory whose names match those in $functions array
Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.ps1 |
	Where-Object { $_.Name -in $functions -and $_.FullName -ne $PSCommandPath } |
	ForEach-Object {
		Write-Output "Importing `e[38;2;0;255;0m$($_.BaseName)`e[0m..."
		. $_.FullName
	}