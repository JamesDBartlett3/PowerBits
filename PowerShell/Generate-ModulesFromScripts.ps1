[PSCustomObject]$ModuleList = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "ModuleList.json") | ConvertFrom-Json

foreach ($module in $ModuleList) {
		$module.name
		foreach ($function in $module.functions) {
				$function
				$script = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "Scripts/$($function).ps1")
				$script | Add-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "$($function).psm1") -WhatIf
		}
}