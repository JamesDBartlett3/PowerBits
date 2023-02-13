Write-Host "Upgrading Enterprise Data Gateway on this server to latest version."
Write-Host "Press any key to continue..."
[void][System.Console]::ReadKey($true)
Login-DataGatewayServiceAccount
Install-DataGateway -AcceptConditions
#Start-Process "C:\Program Files\On-premises data gateway\EnterpriseGatewayConfigurator.exe"
Start-Process "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\On-premises data gateway.lnk"