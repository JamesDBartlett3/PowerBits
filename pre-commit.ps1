#!pwsh
Set-Location $(git rev-parse --show-toplevel)
. './PowerShell/Generate-ModulesFromScripts.ps1'
git add "./PowerShell/Modules/*"
exit $LASTEXITCODE

<#
Contents of pre-commit file in .git/hooks folder:
#!/bin/sh
echo 
exec pwsh.exe -ExecutionPolicy RemoteSigned -File '.\pre-commit.ps1'
exit
#>