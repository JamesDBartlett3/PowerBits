<# 
  .SYNOPSIS
    Title: Import-PSHacks4PowerBIModules
    Author: @JamesDBartlett3@techhub.social (James D. Bartlett III)
  
  .DESCRIPTION
    - Imports modules for the "PowerShell Hacks for Power BI" demo session
  
  .EXAMPLE
    .\Import-PSHacks4PowerBIModules.ps1
  
  .NOTES
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>

# Pre-emptively import problematic modules
Import-Module Az.Resources -Force -ErrorAction SilentlyContinue | Out-Null

# List of modules to import
[PSCustomObject]$ModuleList = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath 'ModuleList.json') | ConvertFrom-Json

# Import all modules in the list
foreach ($module in $ModuleList) {
  $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules/$($module.name).psm1"
  Write-Host "Importing `e[38;2;0;255;0m$($module.name)`e[0m module..."
  Import-Module $modulePath -Force
}
Write-Host "Done."