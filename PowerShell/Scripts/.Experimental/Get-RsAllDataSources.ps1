<#
.SYNOPSIS
Retrieves all data sources from a Reporting Services server and outputs the results to a text file.
 
.DESCRIPTION
This script retrieves all data sources from a Reporting Services server and outputs the results to a text file. 
The script uses the rs.exe utility to connect to the server and retrieve the data sources. 
The results are then parsed to extract the usernames used in the data sources and output to a separate file.

.PARAMETER ServerName (Mandatory)
The name of the Reporting Services server. This should be the server name or IP address. Do not include the protocol or port number.

.PARAMETER PortNumber (Optional)
Port number on which the report server is running. Default is 443.

.PARAMETER OpenOutput (Optional)
Switch parameter to open the output files after the script has completed.

.EXAMPLE
Get-RsAllDataSources -ServerName "reportserver.example.com" -PortNumber 4443 -OpenOutput

This will retrieve all data sources from the report server "reportserver.example.com" on port 4443 and output the results to a text file, then open the file.

.NOTES
  ACKNOWLEDGEMENTS
    - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
    - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>

[CmdletBinding()]
param (
  [Parameter(Mandatory)][string]$ServerName,
  [Parameter()][int]$PortNumber = 443,
  [Parameter()][switch]$OpenOutput
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command "rs.exe" -ErrorAction SilentlyContinue)) {
  Write-Error "Command 'rs.exe' not found. Please install the Report Server Command Prompt Utilities (see: https://learn.microsoft.com/en-us/sql/reporting-services/tools/rs-exe-utility-ssrs)"
  Exit
}

# Set the security protocol to TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$DatePrefix = "$((Get-Date).ToString('yyyyMMdd-HHmm'))"
$OutFile = Join-Path -Path $PSScriptRoot -ChildPath "$($DatePrefix)_$($ServerName)_AllDataSources.txt"
$UserNamesFile = $OutFile.Replace("DataSources", "UserNames")
$UserNames = @()
$RsScriptBody = @"
' Originally from http://social.msdn.microsoft.com/Forums/en/sqlreportingservices/thread/a2a9ebe1-0417-46bf-8589-ae4e4a16181c in a post by Igor Gelin
' Additional fields from http://msdn.microsoft.com/en-us/library/reportservice2005.datasourcedefinition.aspx added.
' This appears to ONLY work for Shared data sources, not Custom data source or embedded data source (i.e. per-report data sources)
Private logFilePath As String = "DataSources.csv"
Private SrcServer As String
Public Sub Main()
Dim pi as Integer = rs.Url.IndexOf("://")
If Not pi = -1 Then
  SrcServer = rs.Url.Substring(pi+3)
End If
SrcServer = SrcServer.Substring(0, SrcServer.IndexOf("/"))
SrcServer = SrcServer.Substring(0, SrcServer.IndexOf(":"))
Dim items As CatalogItem() = Nothing
Dim dataSource As DataSourceDefinition
Dim count as Integer = 0
Try
  items = rs.ListChildren("/", True)
  Console.WriteLine("{0} DataSources:", SrcServer)
  Console.WriteLine()
  Console.WriteLine("===================================")
  For Each catalogItem as CatalogItem in items
  If (catalogItem.Type = ItemTypeEnum.DataSource)
  Console.WriteLine(catalogItem.Path)
  dataSource = rs.GetDataSourceContents(catalogItem.Path)
  If Not (dataSource Is Nothing) Then
  Console.WriteLine(" Connection String: {0}", dataSource.ConnectString)
  Console.WriteLine(" Extension name: {0}", dataSource.Extension)
  Console.WriteLine(" Credential retrieval: {0}", dataSource.CredentialRetrieval)
  Console.WriteLine(" Windows credentials: {0}", dataSource.WindowsCredentials)
  Console.WriteLine(" Username: {0}", dataSource.UserName)
  Console.WriteLine(" Password: {0}", dataSource.Password)
  Console.WriteLine(" Enabled: {0}", dataSource.Enabled)
  Console.WriteLine(" EnabledSpecified: {0}", dataSource.EnabledSpecified)
  Console.WriteLine(" ImpersonateUser: {0}", dataSource.ImpersonateUser)
  Console.WriteLine(" ImpersonateUserSpecified: {0}", dataSource.ImpersonateUserSpecified)
  Console.WriteLine(" OriginalConnectStringExpressionBased: {0}", dataSource.OriginalConnectStringExpressionBased)
  Console.WriteLine(" Prompt: {0}", dataSource.Prompt)
  Console.WriteLine(" UseOriginalConnectString: {0}", dataSource.UseOriginalConnectString)
  Console.WriteLine("===================================")
  End If
  count = count + 1
  End If
  Next catalogItem
  Console.WriteLine()
  Console.WriteLine("Total {0} datasources", count)
  Catch e As IOException
  Console.WriteLine(e.Message)
End Try
End Sub
"@

$RsScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "$($DatePrefix)_$($ServerName)_AllDataSources_rsscript.vb"

Set-Content -Path $RsScriptPath $RsScriptBody

rs.exe -i $RsScriptPath -s "https://$($ServerName):$($PortNumber)/ReportServer" | Out-File $OutFile

Remove-Item $RsScriptPath

Get-Content $OutFile | ForEach-Object {
  if($_ -like " Username:*") {
    $UserNames += $_.replace(" Username: ", "")
  }
}

Set-Content -Path $UserNamesFile -Value "$ServerName DataSource UserNames:`n"
$UserNames | Where-Object {$_.Length -gt 0} | Sort-Object -Unique | Add-Content -Path $UserNamesFile

if ($OpenOutput) {
  $OutFile, $UserNamesFile | Invoke-Item
}