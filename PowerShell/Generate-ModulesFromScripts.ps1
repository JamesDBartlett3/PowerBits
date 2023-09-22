[PSCustomObject]$ModuleList = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "ModuleList.json") | ConvertFrom-Json

foreach ($module in $ModuleList) {
		$module.name
		foreach ($function in $module.functions) {
				"Module: $($module.name) Function: $function"
				$script = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "Scripts/$($function).ps1")
				"Function $function {`n" | Add-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "$($module.name)_test.psm1")
				$script | Add-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "$($module.name)_test.psm1")
				"`n}" | Add-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "$($module.name)_test.psm1")
		}
}