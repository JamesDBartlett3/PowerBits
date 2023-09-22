<# 
	.SYNOPSIS
		Title: Import-PSHacks4PowerBIModules
		Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)

	.DESCRIPTION
		- Imports modules for the "PowerShell Hacks for Power BI" demo session

	.NOTES
		- Version: 1.0

	.EXAMPLE
		. .\Import-PSHacks4PowerBIModules.ps1
#>

# Pre-emptively import problematic modules
Import-Module Az.Resources -Force -ErrorAction SilentlyContinue | Out-Null

# List of modules to import
$modules = @(
	"Get-UserDatasets.psm1"
	, "Join-UserDatasetsWithWorkspaces.psm1"
	, "Update-UserDatasetsOwner.psm1"
	, "Copy-PowerBIReportContentToBlankPBIXFile.psm1"
	, "Get-DataGatewayNodesStatus.psm1"
	, "Export-PowerBIReportsFromWorkspaces.psm1"
  , "Export-PowerBIWorkspacesSecurity.psm1"
	, "BareDatasets.psm1"
	, "Utilities.psm1"
)

# Dotsource all modules in current directory whose names match those in $modules array
Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.psm1 |
	Where-Object { $_.Name -in $modules -and $_.FullName -ne $PSCommandPath } |
	ForEach-Object {
		Write-Host "Importing `e[38;2;0;255;0m$($_.BaseName)`e[0m module..."
		Import-Module $_.FullName -Force
	}
Write-Host "Done."