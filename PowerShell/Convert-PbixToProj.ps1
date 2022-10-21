<#
    .SYNOPSIS
        Title: Convert-PbixToProj
        Author: @JamesDBartlett3 (James D. Bartlett III)

    .DESCRIPTION
        Convert a Power BI report to a pbi-tools project

    .PARAMETER PbixPath
        The path to the PBIX file to convert

    .PARAMETER ExtractPath
        The path where the pbi-tools project will be created

    .PARAMETER ModelSerialization
        The model serialization format to use. Valid values are "Default" and "Raw"

    .PARAMETER MashupSerialization
        The mashup serialization format to use. Valid values are "Default," "Raw," and "Expanded"

    .EXAMPLE
        Convert-PbixToProj -PbixPath ".\MyReport.pbix"

    .EXAMPLE
        Convert-PbixToProj -PbixPath ".\MyReport.pbix" -ExtractPath ".\MyReport" -ModelSerialization "Raw" -MashupSerialization "Expanded"

    .LINK
        https://pbi.tools
        https://pbi.tools/cli/usage.html#extract

    .NOTES
        TODO
            - Write function body
            - Add requirement for pbi-tools to be installed
#>

Function Convert-PbixToProj {

    #Requires -PSEdition Core

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$PbixPath,
        [Parameter(Mandatory = $false)][string]$ExtractPath,
        [Parameter(Mandatory = $false)][ValidateSet("Default", "Raw")][string]$ModelSerialization,
        [Parameter(Mandatory = $false)][ValidateSet("Default", "Raw", "Expanded")][string]$MashupSerialization
    )

    try {
        Invoke-Expression pbi-tools | Out-Null
    }

    catch {
        Write-Error "'pbi-tools' command not found. See: https://pbi.tools"
        Write-Warning $Error[0]
    }

    finally{
        if (!$Error[0]) {
            $ep = if($ExtractPath) {"-extractFolder $ExtractPath"}
            $mod = if($ModelSerialization) {"-modelSerialization $ModelSerialization"}
            $mash = if($MashupSerialization) {"-mashupSerialization $MashupSerialization"}
            $command = "pbi-tools extract -pbixPath $PbixPath $ep $mod $mash"
            Write-Debug "Running command: $command"
            Write-Output "Extracting $PbixPath $(if($ExtractPath) {"to $ExtractPath"})"
            Invoke-Expression $command | Out-Null
        }
        
    }

}