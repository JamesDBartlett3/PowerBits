param(
  [string]$SourceReportServerName,
  [string]$SourceReportServerInstanceName = "ReportServer",
  [int]$SourcePortNumber = 443,
  [string]$SourceReportServerProtocol = "https",
  [string[]]$SourceFoldersToMigrate = "/",
  [string]$DestinationReportServerName,
  [string]$DestinationReportServerInstanceName = "ReportServer",
  [int]$DestinationPortNumber = 443,
  [string]$DestinationReportServerProtocol = "https",
  [string]$DestinationFolder = "/",
  [string]$RsScriptFile = ".\ssrs_migration.vb"
)

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12;

$FoldersToMigrate | ForEach-Object {
  Write-Host "Migrating $_..."
  # TODO: Refactor for safety
  $CommandString = "rs.exe -i ""$RsScriptFile"" -e Mgmt2010 -s ""$($SourceReportServerProtocol)://$SourceReportServerName/$SourceReportServerInstanceName"" -v f=""$_"" -v ts=""$($DestinationReportServerProtocol)://$DestinationReportServerName/$DestinationReportServerInstanceName"" -v tf=""/Legacy"" -v security=""True"" -v unattended=""True"" -v logprefix=""$((Get-Date).ToString('yyyyMMdd'))_"""
  Invoke-Expression -Command $CommandString
}