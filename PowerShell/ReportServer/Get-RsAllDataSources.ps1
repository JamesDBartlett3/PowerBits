param(
  $ServerName = "localhost",
  $ReportServerInstanceName = "ReportServer",
  $PortNumber = 443,
  $DatePrefix = "$((Get-Date).ToString('yyyyMMdd-HHmm'))",
  $OutFile = ".\$($DatePrefix)_$($ServerName)_AllDataSources.txt"
)

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
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

$RsScriptPath = ".\$($DatePrefix)_$($ServerName)_AllDataSources_rsscript.vb"

Set-Content -Path $RsScriptPath $RsScriptBody

rs.exe -i $RsScriptPath -s "https://$($ServerName):$($PortNumber)/$($ReportServerInstanceName)" | Out-File $OutFile

Remove-Item $RsScriptPath

Get-Content $OutFile | ForEach-Object {
  if($_ -like " Username:*") {
    $UserNames += $_.replace(" Username: ", "")
  }
}

Set-Content -Path $UserNamesFile -Value "$ServerName DataSource UserNames:`n"
$UserNames | Where-Object {$_.Length -gt 0} | Sort-Object -Unique | Add-Content -Path $UserNamesFile

$OutFile, $UserNamesFile | Invoke-Item