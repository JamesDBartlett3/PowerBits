# Credit: @ruiromano on GitHub
# Source: https://github.com/RuiRomano/pbimonitor

[CmdletBinding()]
Param(
	[Parameter(Mandatory)][ScriptBlock]$script,
	[Parameter][int]$sleepSeconds = 3601,
	[Parameter][int]$tentatives = $null
)
try {
	Invoke-Command -ScriptBlock $script
} catch {
	$ex = $_.Exception
	$errorText = $ex.ToString()
	if ($errorText -like '*HttpRequestException*' -and $errorText -like '*429*') {
		Write-Host "'429 (Too Many Requests)' Error - Sleeping for $sleepSeconds seconds before trying again..." -ForegroundColor Yellow
		if ($tentatives) {$tentatives = $tentatives - 1}
		if ($tentatives -eq 0) {
			throw '[Wait-On429Error] Max tentatives reached!'
		} else {
			Start-Sleep -Seconds $sleepSeconds
			Wait-On429Error -script $script -sleepSeconds $sleepSeconds -tentatives $tentatives
		}
	} else {
		throw
	}
}