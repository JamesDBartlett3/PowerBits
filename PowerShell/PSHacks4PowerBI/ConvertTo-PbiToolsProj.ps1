Function ConvertTo-PbiToolsProj {
    <#
    .SYNOPSIS
        Title: ConvertTo-PbiToolsProj
        Author: @JamesDBartlett3 (James D. Bartlett III)

    .DESCRIPTION


    .PARAMETERS
        -

    .EXAMPLE
        ConvertTo-PbiToolsProj -PbixPath ".\MyReport.pbix"
        ConvertTo-PbiToolsProj -PbixPath ".\MyReport.pbix" -ExtractFolder ".\MyReport" -ModelSerialization "Raw" -MashupSerialization "Expanded"

    .TODO
        - Write function body

    #>

    #Requires -PSEdition Core
    #Requires 

        Param(
        [parameter(Mandatory = $true)][string]$PbixPath,
        [parameter(Mandatory = $false)][string]$ExtractFolder = $null,
        [parameter(Mandatory = $false)][string]$ModelSerialization = "Default",
        [parameter(Mandatory = $false)][string]$MashupSerialization = "Default"
    )

}