<#
  .SYNOPSIS
    Generates PowerShell modules from scripts.

  .DESCRIPTION
    This script requires:
      - a file named ModuleList.json in the same directory as itself.
      - a folder named Scripts in the same directory as itself.
      - a folder named Modules in the same directory as itself.
      - one or more PowerShell script files in the Scripts folder, each with the same name as a function listed in the ModuleList.json file.

    The ModuleList.json file should contain an array of objects with the following properties:
      - name: The name of the module to create
      - functions: An array of function names to include in the module

    This script will use the ModuleList.json file as a guide to generate one or more modules, each containing one or more functions.
    The generated modules and the functions they contain will be named according to the ModuleList.json file, and saved in the Modules folder.

    After generating each module, this script will also follow the rules defined in the $formatterSettings variable to format the module's contents
    using the PSScriptAnalyzer module's Invoke-Formatter cmdlet. This is done to ensure that the generated modules are formatted consistently.

  .LINK
    [Source code](https://github.com/JamesDBartlett3/PowerBits)

  .LINK
    [The author's blog](https://datavolume.xyz)

  .LINK
    [Follow the author on LinkedIn](https://www.linkedin.com/in/jamesdbartlett3/)

  .LINK
    [Follow the author on Mastodon](https://techhub.social/@JamesDBartlett3)

  .LINK
    [Follow the author on BlueSky](https://bsky.app/profile/jamesdbartlett3.bsky.social)

  .NOTES
    TODO:
      - If a script has been modified but not staged, get its contents from the previous commit.
      - Handle deleted scripts.
      - Once the above is implemented, remove the -GenerateAll switch from pre-commit.ps1.
    ACKNOWLEDGEMENTS
      - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
      - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>

Param(
  [switch]$GenerateAll
)

#Requires -Modules PSScriptAnalyzer

# Define the settings to use when formatting the generated modules
$formatterSettings = @{
  IncludeRules = @('PSPlaceOpenBrace', 'PSUseConsistentIndentation')
  Rules        = @{
    PSPlaceOpenBrace           = @{
      Enable     = $true
      OnSameLine = $true
    }
    PSUseConsistentIndentation = @{
      Enable          = $true
      IndentationSize = 2
    }
  }
}

# Create a PSCustomObject from the ModuleList.json file
[PSCustomObject]$ModuleList = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath 'ModuleList.json') | ConvertFrom-Json

# Get a list of all files with staged changes
$stagedFiles = (git diff --cached --name-only).ForEach({ [system.io.path]::GetFileNameWithoutExtension($_) })

# Add a property to each module object to indicate whether it should be updated
# based on whether any of its functions have changed since the last commit
$ModuleList.ForEach({
    $_ | Add-Member -NotePropertyName "update" -NotePropertyValue $_.functions.ForEach({
        $_ -in $stagedFiles
      }).Contains($true)
  })

$ModuleList = $ModuleList | Where-Object { $_.update -or $GenerateAll }

foreach ($module in ($ModuleList)) {
  $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules/$($module.name).psm1"
  Write-Verbose "Module Name: $($module.name) -- Path: $modulePath"
  $moduleContent = '#Requires -PSEdition Core'
  foreach ($function in $module.functions) {
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Scripts/$($function).ps1"
    Write-Verbose "Function: $function -- Script Path: $scriptPath"
    $moduleContent += "`nFunction $function {"
    $moduleContent += (Get-Content -Raw -Path $scriptPath).Replace('#Requires -PSEdition Core', '')
    $moduleContent += "`n}"
  }
  # Replace multiple newlines and/or carriage returns with one newline, then remove any lines that only contain whitespace
  $moduleContent = $moduleContent -Replace "`r+", "`n" -Replace "`n+", "`n" -Replace "`n\s+`n", "`n"
  Invoke-Formatter -ScriptDefinition $moduleContent -Settings $formatterSettings | Set-Content -Path $modulePath -Force
}