#Requires -Modules PSScriptAnalyzer

[PSCustomObject]$ModuleList = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "ModuleList.json") | ConvertFrom-Json

foreach ($module in $ModuleList) {
	$script = @'
	#Requires -PSEdition Core
'@
	foreach ($function in $module.functions) {
		"Module: $($module.name) Function: $function"
		$script += "`nFunction $function {"
		$script += (Get-Content -Raw -Path (Join-Path -Path $PSScriptRoot -ChildPath "Scripts/$($function).ps1")).Replace("#Requires -PSEdition Core", "")
		$script += "`n}"
	}
	Invoke-Formatter -ScriptDefinition $script.Replace("`r`n","`n") | Set-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "$($module.name).psm1")
}