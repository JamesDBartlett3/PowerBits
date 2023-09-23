#!pwsh
Set-Location $(git rev-parse --show-toplevel)
. './PowerShell/Generate-ModulesFromScripts.ps1'
git add "./PowerShell/Modules/*"
exit $LASTEXITCODE