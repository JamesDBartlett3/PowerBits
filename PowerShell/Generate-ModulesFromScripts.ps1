#Requires -Modules PSScriptAnalyzer

[PSCustomObject]$ModuleList = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath 'ModuleList.json') | ConvertFrom-Json

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

foreach ($module in $ModuleList) {
	$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules/$($module.name).psm1"
	Write-Verbose "Module Name: $($module.name) -- Path: $modulePath"
	$moduleContent = '#Requires -PSEdition Core'
	foreach ($function in $module.functions) {
		$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Scripts/$($function).ps1"
		Write-Verbose "Function: $function -- Script Path: $scriptPath"
		$moduleContent += "`nFunction $function {"
		$moduleContent += (Get-Content -Raw -Path $scriptPath).Replace('#Requires -PSEdition Core', '').Trim() -ne ''
		$moduleContent += "`n}"
	}
	# Replace multiple newlines and/or carriage returns with one newline
	# $moduleContent = $moduleContent.Replace("`r+", "`n").Replace("`n+", "`n") -Replace "`n\s+`n", "`n" -Replace "`n\t+`n", "`n"
	Invoke-Formatter -ScriptDefinition $moduleContent -Settings $formatterSettings | Set-Content -Path $modulePath -Force
}