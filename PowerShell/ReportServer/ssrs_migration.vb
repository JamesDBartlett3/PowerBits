' This is a Microsoft SQL Server Reporting Services RSS script that migrates content from one Reporting Services server to another.
' Run the script with the rs.exe utility.  Rs.exe is installed by Reporting Services.
' For details about rs.exe see: https://docs.microsoft.com/sql/reporting-services/tools/rs-exe-utility-ssrs
'
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' For more detailed instructions and examples of using the script, see: https://docs.microsoft.com/sql/reporting-services/tools/sample-reporting-services-rs-exe-script-to-copy-content-between-report-servers
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
' The script supports report server versions SQL Server 2008 R2 and later and Power BI Report Server.
' The script supports both native mode report servers and SharePoint mode report servers.
' The script can be run from the source or target server.
'
'-----------------------------------------------------------------------------------------------------------
'
' To use the script:
' 1) Download ssrs_migration.rss
' 2) Open a command prompt and navigate to the folder containing ssrs_migration.rss, for example c:\rss
' 3) Run the following command (in one line):
'			rs.exe -i ssrs_migration.rss -e Mgmt2010
'
' 		-s SOURCE_URL						                  'URL of the source RS server.
'			-u domain\username -p password	          'Credentials for source server. OPTIONAL, default credentials are used if missing.
'			-v st="SITE"						                  'Specifies SharePoint site, in case source server is in SharePoint integrated mode
'			-v f="SOURCEFOLDER"					              'Set to "/" for migrating everything, or to something like "/folder/subfolder" for partial migration. Everything within this folder will be copied. OPTIONAL, default is "/".
'			-v ts="TARGET_URL" 					              'URL of the target RS server"
'			-v tu="domain\username" -v tp="password"	'Credentials for target server. OPTIONAL, default credentials are used if missing.
'			-v tst="SITE"						                  'Specifies SharePoint site, in case target server is in SharePoint integrated mode
'			-v tf="TARGETFOLDER"				              'Set to "/" for migrating into the root level, or to something like "/folder/subfolder" for copying into some folder, which must be already existing. Everything within "SOURCEFOLDER" will be copied into "TARGETFOLDER". OPTIONAL, default is "/".
'			-v security= "True/False"			            'If set to "False", destination catalog items will inherit security setting according to the settings of the target system. Default is false.
'     -v logprefix="PREFIX"                     'Prefix name to the output log file.
'			-v unattended="True/False"					      'Run without asking for confirmation
'			
' Example: rs.exe -i ssrs_migration.rss -e Mgmt2010 -s http://server1/reportserver -v ts="http://server2/_vti_bin/reportserver"
' Example: rs.exe -i ssrs_migration.rss -e Mgmt2010 -s http://server1/reportserver -u domain1\user1 -p password1 -v f="/SOURCEFOLDER" -v ts="http://server2/_vti_bin/reportserver" -v tu="domain1\user2" -v tp="password2" -v tf="/TARGETFOLDER"
'
' SOURCE_URL and TARGET_URL must be valid reportserver URLs pointing to the source and target RS report server.
' In native mode a RS report server URL looks like: http://servername/reportserver
' In SharePoint integrated mode such a URL looks like: http://servername/_vti_bin/reportserver
'
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' For more detailed instructions and examples of using the script, see: https://docs.microsoft.com/sql/reporting-services/tools/sample-reporting-services-rs-exe-script-to-copy-content-between-report-servers
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>'
'
' LIMITATIONS:
' - Passwords are not migrated, and must be re-entered (e.g. data sources with stored credentials)
'
' ADDITIONAL INFO:
' The virtual folder structure presented to the user in SharePoint might be different
' than the physical structure, which is used for by this script.
' Open http://servername/_vti_bin/reportserver in a browser to see the non-virtual folder structure.
' This is helpful for setting SrcFolder and SnkFolder to something other than "/" for a server in SharePoint integrated mode.

Private RsSrc As ReportingService2010
Private SrcProtocol As String = "http"
Private SrcServer As String
Private SrcIsNative As Boolean
Private SrcSite As String = ""
Private SrcFolder As String = "/"

Private RsSnk As ReportingService2010
Private SnkProtocol As String = "http"
Private SnkServer As String
Private SnkIsNative As Boolean
Private SnkSite As String = ""
Private SnkFolder As String = "/"

Private MigrateSecurity As Boolean = True

Private srcItems() As CatalogItem = Nothing
Private snkItems() As CatalogItem = Nothing
Private schedules As Schedule() = Nothing
Private roles As Role() = Nothing
Private policies As Policy() = Nothing

Private srcSiteUrl As String = Nothing
Private snkSiteUrl As String = Nothing

Private logFilePath As String = "MigrationLog.csv"
Private RunUnattended As Boolean = False

''''''''''''''''''''''''''''''''''''''''''''''
Sub Main()

  ' Initialize variables
  RsSrc = rs
  Dim pi as Integer = rs.Url.IndexOf("://")
  If Not pi = -1 Then
  SrcProtocol = rs.Url.Substring(0, pi)
  SrcServer = rs.Url.Substring(pi+3)
  End If
  SrcServer = SrcServer.Substring(0, SrcServer.IndexOf("/"))
  SrcIsNative = Not rs.Url.Contains("/_vti_bin")
  
  If Not Me.GetType().GetField("st") Is Nothing Then
  SrcSite = Me.GetType().GetField("st").GetValue(Me)
  End If	
  If SrcSite = "/" Then
  SrcSite = ""
  End If
  If Not SrcSite.StartsWith("/") And Not SrcSite = "" Then
  SrcSite = "/" + SrcSite
  End If
  
  If Not Me.GetType().GetField("f") Is Nothing Then
  SrcFolder = Me.GetType().GetField("f").GetValue(Me)
  End If	
  If Not SrcFolder.StartsWith("/") And Not SrcFolder = "/" Then
  SrcFolder = "/" + SrcFolder
  End If
  
  SrcFolder = SrcSite + SrcFolder
  
  RsSnk = New ReportingService2010()
  If Me.GetType().GetField("tu") Is Nothing Then
  RsSnk.Credentials = System.Net.CredentialCache.DefaultCredentials
  Else
  Dim user As String = Me.GetType().GetField("tu").GetValue(Me)
  Dim password As String = Me.GetType().GetField("tp").GetValue(Me)

  If user.contains("\") Then			
    Dim domainuser As String() = user.Split(New [Char]() {"\"c})
    RsSnk.Credentials = New System.Net.NetworkCredential(domainuser(1), password, domainuser(0))	
  Else
    RsSnk.Credentials = New System.Net.NetworkCredential(user, password)		
  End If
  End If
  RsSnk.Url = ts + "/reportservice2010.asmx"

  pi = RsSnk.Url.IndexOf("://")
  If Not pi = -1 Then
  SnkProtocol = RsSnk.Url.Substring(0, pi)
  SnkServer = RsSnk.Url.Substring(pi+3)		
  End If
  SnkServer = SnkServer.Substring(0, SnkServer.IndexOf("/"))
  SnkIsNative = Not RsSnk.Url.Contains("/_vti_bin")
  
  If Not Me.GetType().GetField("tst") Is Nothing Then
  SnkSite = Me.GetType().GetField("tst").GetValue(Me)
  End If	
  If SnkSite = "/" Then
  SnkSite = ""
  End If
  If Not SnkSite.StartsWith("/") And Not SnkSite = "" Then
  SnkSite = "/" + SnkSite
  End If
  
  If Not Me.GetType().GetField("tf") Is Nothing Then
  SnkFolder = Me.GetType().GetField("tf").GetValue(Me)
  End If	
  If Not SnkFolder.StartsWith("/") And Not SnkFolder = "/" Then
  SnkFolder = "/" + SnkFolder
  End If
  
  If Not Me.GetType().GetField("security") Is Nothing Then
  MigrateSecurity = (Me.GetType().GetField("security").GetValue(Me).ToLower = "true")
  End If

  ' Update the migration log path to prefix if specified
  If Not Me.GetType().GetField("logprefix") Is Nothing Then
  logFilePath = Me.GetType().GetField("logprefix").GetValue(Me) + "_" + logFilePath
  End If

  ' Set the RunUnattended boolean variable from user input
  If Not Me.GetType().GetField("unattended") Is Nothing Then
  RunUnattended = (Me.GetType().GetField("unattended").GetValue(Me).ToLower = "true")
  End If
  
  SnkFolder = SnkSite + SnkFolder
  
  srcSiteUrl = IIf(SrcIsNative, Nothing, SrcProtocol + "://" + SrcServer + SrcSite)
  snkSiteUrl = IIf(SnkIsNative, Nothing, SnkProtocol + "://" + SnkServer + SnkSite)
  
  Console.Write("Retrieve and report the list of items that will be migrated.")
  Console.ForegroundColor = ConsoleColor.Green
  Console.WriteLine("You can cancel the script after step 1 if you do not want to start the actual migration.")
  Console.ResetColor()
  
  If MigrateSecurity And SrcIsNative And SnkIsNative Then	 'Roles and Policies are only migrated from native to native
  Console.WriteLine(Environment.NewLine + "Retrieving roles: ")
  roles = RsSrc.ListRoles("All", srcSiteUrl)
  For Each r As Role In roles
    Console.ForegroundColor = ConsoleColor.DarkGray
    Console.WriteLine("Role: " + r.Name)
    Console.ResetColor()
  Next
  
  Console.WriteLine(Environment.NewLine + "Retrieving system policies: ")
  policies = RsSrc.GetSystemPolicies()
  For each p as Policy in policies
    Console.ForegroundColor = ConsoleColor.DarkGray
    Console.WriteLine("System policy: " + p.GroupUserName)
    Console.ResetColor()
  Next
  End If
  
  Console.WriteLine(Environment.NewLine + "Retrieving schedules: ")
  schedules = RsSrc.ListSchedules(srcSiteUrl)
  For Each s As Schedule In schedules
  Console.ForegroundColor = ConsoleColor.DarkGray
  Console.WriteLine("Schedule: " + s.Name)
  Console.ResetColor()
  Next
  
  Console.WriteLine(Environment.NewLine + "Retrieving catalog items. This may take a while.")
  Dim timeout As Integer = RsSrc.Timeout
  RsSrc.Timeout = 600000 '10 minutes
  srcItems = RsSrc.ListChildren(GetSrcFolderPath(), True) 'Possible catalog item types: Model, Dataset, Component, Resource, DataSource, Folder, Report
  Array.Sort(srcItems, New Comparison(Of CatalogItem)(AddressOf SortParentFoldersFirst))
  RsSrc.Timeout = timeout
  
  For Each ci As CatalogItem In srcItems
  If SrcIsNative And ci.Path.Contains("Users Folders") Then
    Continue For
  End If
  Console.ForegroundColor = ConsoleColor.DarkGray
  Console.WriteLine(ci.TypeName + ": " + ci.Path)	
  Console.ResetColor()
  Next
  
  Console.WriteLine(Environment.NewLine + "--Migration Source--")
  Console.WriteLine("Protocol: " + SrcProtocol)
  Console.WriteLine("Server name: " + SrcServer)
  Console.WriteLine("Mode: " + IIf(SrcIsNative, "Native", "SharePoint integrated"))
  Console.WriteLine("Location: " + SrcFolder)
  Console.WriteLine("")
  Console.WriteLine("--Migration Target--")
  Console.WriteLine("Protocol: " + SnkProtocol)
  Console.WriteLine("Server name: " + SnkServer)
  Console.WriteLine("Mode: " + IIf(SnkIsNative, "Native", "SharePoint integrated"))
  Console.WriteLine("Location: " + SnkFolder)
  
  If Not RunUnattended Then
  Console.ForegroundColor = ConsoleColor.Green
  Console.WriteLine(Environment.NewLine + "Press <Enter> to start migration of items listed above or <Ctrl>+<C> to cancel.")
  Console.ResetColor()
  Console.ReadLine()
  End If

  ' Create destination folder if not present
  Try 
    If SnkFolder.Contains("/") And Not SnkFolder = "/" Then
      Dim targetFolders As String() = SnkFolder.Substring(1).Split("/"c)
      Dim parentFolder As String = "/"

      For Each folderName As String In targetFolders
        Try
          RsSnk.CreateFolder(folderName, parentFolder, Nothing)
        Catch er As Exception
          If (er.Message.Contains("Microsoft.ReportingServices.Diagnostics.Utilities.ItemAlreadyExistsException"))
            Console.ForegroundColor = ConsoleColor.Gray
            Console.WriteLine("Folder already exists: " + folderName)
            Console.ResetColor()
          Else
            Throw ' Unknown Exception
          End If
        End Try

        If parentFolder <> "/" Then
          parentFolder = parentFolder + "/" + folderName
        Else 
          parentFolder = "/" + folderName
        End If
      Next            
    End If
  Catch ex As Exception 
    Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(ex.Message + Environment.NewLine)
    Console.ResetColor()
  End Try

  Try
  'Roles & Policies
  MigrateRoles()				
  MigrateSystemPolicies()
  
  'Schedules
  MigrateSchedules()	
  
  If srcItems Is Nothing OrElse srcItems.Length = 0 Then
    Console.WriteLine("No catalog items were retrieved. Nothing more to do here.")
    Return
  End If							
  
  snkItems = New CatalogItem(srcItems.Length - 1) {}

  'Migrate Catalog Items
  For i As Integer = 0 To srcItems.Length - 1	
    If SrcIsNative And srcItems(i).Path.Contains("Users Folders") Then
    Continue For
    End If
  
    If srcItems(i).TypeName = "Folder" Then
    snkItems(i) = MigrateFolder(srcItems(i))
    ElseIf srcItems(i).TypeName = "DataSource" Then
    snkItems(i) = MigrateDataSource(srcItems(i))
    ElseIf Not srcItems(i).TypeName = "LinkedReport" Then
    snkItems(i) = MigrateCatalogItem(srcItems(i))
    End If
  Next
  
  'Relink Catalog Items (only applicable for (Linked) Reports, Models and Data Sources)
  For i As Integer = 0 To srcItems.Length - 1
    If SrcIsNative And srcItems(i).Path.Contains("Users Folders") Then
    Continue For
    End If
  
    If Not snkItems(i) Is Nothing And (srcItems(i).TypeName = "Model" Or srcItems(i).TypeName = "Report") Then
    RelinkDataSources(snkItems(i), srcItems(i))
    End If

    If Not snkItems(i) Is Nothing And (srcItems(i).TypeName = "Report" Or srcItems(i).TypeName = "DataSet") Then
    RelinkItemReferences(snkItems(i), srcItems(i))
    End If

    If srcItems(i).TypeName = "LinkedReport" And SnkIsNative Then
    snkItems(i) = MigrateLinkedReport(srcItems(i)) 'Migrating a linked report is like relinking
    End If
  Next

  'Migrate items' artefacts
  For i As Integer = 0 To srcItems.Length - 1
    If SrcIsNative And srcItems(i).Path.Contains("Users Folders") Then
    Continue For
    End If
    If (srcItems(i).TypeName = "Report" Or srcItems(i).TypeName = "LinkedReport") And Not snkItems(i) Is Nothing Then
    MigrateReportParameters(snkItems(i), srcItems(i))
    MigrateReportSubscriptions(snkItems(i), srcItems(i))
    MigrateReportHistorySettings(snkItems(i), srcItems(i))
    MigrateReportExecutionOptions(snkItems(i), srcItems(i))
    End If

    If (srcItems(i).TypeName = "Report" Or srcItems(i).TypeName = "LinkedReport" Or srcItems(i).TypeName = "DataSet") And Not snkItems(i) Is Nothing Then

    MigrateReportCacheOptions(snkItems(i), srcItems(i))
    MigrateCacheRefreshPlans(snkItems(i), srcItems(i))
    End If

    If MigrateSecurity And Not snkItems(i) Is Nothing Then
    MigrateItemPolicies(snkItems(i), srcItems(i))
    End If
  Next

  Console.WriteLine("--- END OF OUTPUT ---")
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("GENERAL", "Exception", ex.Message)
  End Try
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''' Migrate global settings'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub MigrateSchedules()
  If schedules Is Nothing OrElse schedules.Length = 0 Then
  Return
  End If
  
  Console.WriteLine("Migrating schedules.")
  
  Try		
  For Each s As Schedule In schedules
    Try		
    Console.Write("Migrating schedule: " + s.Name)
    RsSnk.CreateSchedule(s.Name, s.Definition, snkSiteUrl)
    WriteSuccess()
    Catch e As Exception
    Console.ForegroundColor = ConsoleColor.Red
    Console.WriteLine(" ... FAILURE:")
    Console.WriteLine(e.Message + Environment.NewLine)
    Console.ResetColor()
        LogErrorToFile("SCHEDULES", s.Name, e.Message)
    End Try
  Next

  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("SCHEDULES", "Exception", ex.Message)
  End Try
End Sub

Sub MigrateRoles()

  If roles Is Nothing OrElse roles.Length = 0 Then
  Return
  End If
  
  Console.WriteLine("Migrating roles.")
  
  Try
  For Each r As Role In roles
    Try
    Console.Write("Migrating role: " + r.Name)

    Dim description As String = Nothing
    Dim taskIds As String() = RsSrc.GetRoleProperties(r.Name, srcSiteUrl, description)
    'According to online description this should rather be
    'Dim taskIds As String() = Nothing
    'Dim description As String = RsSrc.GetRoleProperties(r.Name, srcSiteUrl, taskIds)
    'see: http://msdn.microsoft.com/en-us/library/reportservice2010.reportingservice2010.getroleproperties.aspx

    RsSnk.CreateRole(r.Name, description, taskIds)
    WriteSuccess()
    Catch e As Exception
    Console.ForegroundColor = ConsoleColor.Red
    Console.WriteLine(" ... FAILURE:")
    Console.WriteLine(e.Message + Environment.NewLine)
    Console.ResetColor()
        LogErrorToFile("ROLES", r.Name, e.Message)
    End Try
  Next

  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("ROLES", "Exception", ex.Message)
  End Try
End Sub

Sub MigrateSystemPolicies()

  If policies Is Nothing Then
  Return
  End If
  
  Console.Write("Migrating system policies")
  Try		
  RsSnk.SetSystemPolicies(policies)
  WriteSuccess()

  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine("... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("POLICIES", "Exception", ex.Message)
  End Try
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''' Migrate Catalog Items'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function MigrateFolder(srcItem As CatalogItem) As CatalogItem

  Console.Write("Migrating " + srcItem.TypeName + ": " + srcItem.Path)
  Dim result As CatalogItem = Nothing
  Try
  Dim path As String = GetSnkPath(srcItem.Path.Substring(0, srcItem.Path.LastIndexOf("/")))
  result = RsSnk.CreateFolder(srcItem.Name, path, Nothing)
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("FOLDER", srcItem.Path, ex.Message)
  End Try
  Return result
End Function

Function MigrateDataSource(srcItem As CatalogItem) As CatalogItem

  Console.Write("Migrating " + srcItem.TypeName + ": " + srcItem.Path)
  Dim result As CatalogItem = Nothing
  Try
  Dim definition As DataSourceDefinition = RsSrc.GetDataSourceContents(srcItem.Path)
  Dim Name As String = GetSnkFilename(srcItem.Name, srcItem.TypeName)
  Dim Parent As String = GetSnkPath(srcItem.Path.Substring(0, srcItem.Path.LastIndexOf("/")))

  result = RsSnk.CreateDataSource(Name, Parent, False, definition, Nothing)
  WriteSuccess()
  
  If definition.CredentialRetrieval = CredentialRetrievalEnum.Store Then
    Console.ForegroundColor = ConsoleColor.Yellow
      Dim warningMsg as String = "You need to re-enter the password for user '" + definition.UserName + "' in data source '" + result.Path + "'"
    Console.WriteLine(warningMsg)
    Console.ResetColor()

      LogErrorToFile("CATALOGITEM", "Warning: " + result.Path, warningMsg)

  End If

  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("DATASOURCE", srcItem.Path, ex.Message)
  End Try
  Return result
End Function

Function MigrateLinkedReport(srcItem As CatalogItem) As CatalogItem

  Console.Write("Migrating " + srcItem.TypeName + ": " + srcItem.Path)
  Dim result As CatalogItem = Nothing
  Try
  Dim name As String = GetSnkFilename(srcItem.Name, srcItem.TypeName)
  Dim parent As String = GetSnkPath(srcItem.Path.Substring(0, srcItem.Path.LastIndexOf("/")))
  Dim link As String = RsSrc.GetItemLink(srcItem.Path)

  RsSnk.CreateLinkedItem(name, parent, GetSnkPath(link), Nothing)

  result = New CatalogItem
  result.CreatedBy = srcItem.CreatedBy
  result.CreationDate = srcItem.CreationDate
  result.CreationDateSpecified = srcItem.CreationDateSpecified
  result.Description = srcItem.Description
  result.Hidden = srcItem.Hidden
  result.HiddenSpecified = srcItem.HiddenSpecified
  result.ID = srcItem.ID
  result.ItemMetadata = srcItem.ItemMetadata
  result.ModifiedBy = srcItem.ModifiedBy
  result.ModifiedDate = srcItem.ModifiedDate
  result.ModifiedDateSpecified = srcItem.ModifiedDateSpecified
  result.Name = name
  result.Path = IIf(parent = "/", "", parent) + "/" + name
  result.Size = srcItem.Size
  result.SizeSpecified = srcItem.SizeSpecified
  result.TypeName = srcItem.TypeName
  result.VirtualPath = Nothing
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("LINKEDREPORT", srcItem.Path, ex.Message)
  End Try
  Return result
End Function

Function MigrateCatalogItem(srcItem As CatalogItem) As CatalogItem

  Console.Write("Migrating " + srcItem.TypeName + ": " + srcItem.Path)
  Dim result As CatalogItem = Nothing
  Try
  Dim definition As Byte() = RsSrc.GetItemDefinition(srcItem.Path)

  Dim propertiesTemplate() As [Property] = New [Property](3) {}
  propertiesTemplate(0) = New [Property]
  propertiesTemplate(0).Name = "MIMEType"
  propertiesTemplate(1) = New [Property]
  propertiesTemplate(1).Name = "ReportTimeout"
  propertiesTemplate(2) = New [Property]
  propertiesTemplate(2).Name = "Hidden"
  propertiesTemplate(3) = New [Property]
  propertiesTemplate(3).Name = "Description"

  Dim properties As [Property]() = RsSrc.GetProperties(srcItem.Path, propertiesTemplate)

  Dim warnings() As Warning = Nothing

  result = RsSnk.CreateCatalogItem(srcItem.TypeName, GetSnkFilename(srcItem.Name, srcItem.TypeName), GetSnkPath(srcItem.Path.Substring(0, srcItem.Path.LastIndexOf("/"))), False, definition, properties, warnings)
  WriteSuccess()
  
  If result.TypeName <> "Component" And result.TypeName <> "Resource" Then
    Dim datasources As DataSource() = RsSrc.GetItemDataSources(srcItem.Path)
    For Each d As DataSource in datasources
    If TypeOf d.Item Is DataSourceDefinition Then
      Dim dsd As DataSourceDefinition = d.Item
      If dsd.CredentialRetrieval = CredentialRetrievalEnum.Store Then
      Console.ForegroundColor = ConsoleColor.Yellow
            Dim warningMsg as String = "You need to re-enter the password for user '" + dsd.UserName + "' in data source '" + d.Name + "' contained in " + result.TypeName + " '" + result.Path + "'"
      Console.WriteLine(warningMsg)
      Console.ResetColor()

            LogErrorToFile("CATALOGITEM", "Warning: " + result.TypeName, warningMsg)
      End If
    End If		
    Next
  End If

  If Not warnings Is Nothing Then
    Console.ForegroundColor = ConsoleColor.Yellow
    Console.WriteLine("")
    For Each w As Warning In warnings
    If w.Code <> "rsDataSourceReferenceNotPublished" And w.Code <> "rsDataSetReferenceNotPublished" Then	'This should be fine after re-linking
      Console.WriteLine("Warning: " + w.Message)
      LogErrorToFile("CATALOGITEM", "Warning: " + srcItem.Path, w.Message)
    End If
    Next
    Console.ResetColor()
  Else
    Console.WriteLine("")
  End If

  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(Environment.NewLine + ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("CATALOGITEM", srcItem.Path, ex.Message)
  End Try
  Return result
End Function


'Re-link DataSources of a Report or a Model
Sub RelinkDataSources(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Re-linking data source for item " + snkReport.Name)
  Try
  Dim srcDataSources As DataSource() = RsSrc.GetItemDataSources(srcReport.Path)
  If srcDataSources Is Nothing Then
    Return
  End If

  Dim snkDataSources As System.Collections.Generic.List(Of DataSource) = New System.Collections.Generic.List(Of DataSource)

  For Each srcDS As DataSource In srcDataSources

    If TypeOf srcDS.Item Is DataSourceReference And Not srcReport.TypeName = "Report" Then 'Report references are migrated with RelinkItemReferences

    Dim snkDSRef As DataSourceReference = New DataSourceReference
    Dim srcDSRef As DataSourceReference = srcDS.Item

    snkDSRef.Reference = GetSnkFilename(GetSnkPath(srcDSRef.Reference), "DataSource")
    Dim snkDS As DataSource = New DataSource
    snkDS.Name = srcDS.Name
    snkDS.Item = snkDSRef
    snkDataSources.Add(snkDS)

    ElseIf TypeOf srcDS.Item Is DataSourceDefinition Then
    snkDataSources.Add(srcDS)
    End If
  Next

  If snkDataSources.Count > 0 Then
    RsSnk.SetItemDataSources(snkReport.Path, snkDataSources.ToArray)
  End If
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("RELINKDATASOURCE", snkReport.Name, ex.Message)
  End Try
End Sub

'Re-link items's references
Sub RelinkItemReferences(snkItem As CatalogItem, srcItem As CatalogItem)
  Console.Write("Re-linking references for item " + snkItem.Name)

  Try
  Dim srcReferences As ItemReferenceData() = RsSrc.GetItemReferences(srcItem.Path, Nothing)

  If srcReferences Is Nothing Then
    Return
  End If

  Dim snkReferences As System.Collections.Generic.List(Of ItemReference) = New System.Collections.Generic.List(Of ItemReference)

  For Each srcRef As ItemReferenceData In srcReferences
    Dim type As String = RsSrc.GetItemType(srcRef.Reference)
    Dim snkRef As ItemReference = New ItemReference
    snkRef.Name = srcRef.Name
    snkRef.Reference = GetSnkFilename(GetSnkPathRef(srcRef.Reference), type)

    snkReferences.Add(snkRef)
  Next

  If snkReferences.Count > 0 Then
    RsSnk.SetItemReferences(snkItem.Path, snkReferences.ToArray)
  End If

  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("RELINKREFERENCES", snkItem.Name, ex.Message)
  End Try
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''' Migrate report artefacts ''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Migrate report's subscriptions
Sub MigrateReportSubscriptions(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Migrating subscriptions for report " + snkReport.Name + ": ")

  Try
  Dim srcSubscriptions As Subscription() = RsSrc.ListSubscriptions(srcReport.Path)

  If srcSubscriptions Is Nothing OrElse srcSubscriptions.Length = 0 Then
    Console.WriteLine("0 items found.")
    Return
  End If
  Console.WriteLine(srcSubscriptions.Length.ToString + " items found.")
  For Each srcSub As Subscription In srcSubscriptions

    Dim extensionSettings As ExtensionSettings = Nothing
    Dim description As String = ""
    Dim active As ActiveState = Nothing
    Dim status As String = Nothing
    Dim eventType As String = Nothing
    Dim matchData As String = Nothing
    Dim parameters As ParameterValueOrFieldReference() = Nothing

    If srcSub.IsDataDriven Then
    Console.Write("Migrating data-driven subscription ")
    Dim dataRetrievalPlan As DataRetrievalPlan = Nothing
    RsSrc.GetDataDrivenSubscriptionProperties(srcSub.SubscriptionID, extensionSettings, dataRetrievalPlan, description, active, status, eventType, matchData, parameters)				
    If extensionSettings.Extension = "Report Server FileShare" Then
      Dim newParameterValues() As [ParameterValueOrFieldReference] = New [ParameterValueOrFieldReference](extensionSettings.ParameterValues.Length) {}
      Dim passwordParamFound As Boolean = False
      Array.Copy(extensionSettings.ParameterValues, newParameterValues, extensionSettings.ParameterValues.Length)
      For Each pvofr As ParameterValueOrFieldReference In extensionSettings.ParameterValues
      If TypeOf pvofr Is ParameterValue Then
        Dim pv As ParameterValue = pvofr
        If pv.Name.Contains("PASSWORD") Then
        'Subscription has Password field linked to dataset field
        passwordParamFound = True
        Exit For
        End If
      Else If TypeOf pvofr Is ParameterFieldReference Then
        Dim pfr As ParameterFieldReference = pvofr
        If pfr.ParameterName.Contains("PASSWORD") Then
        'Subscription has Password field linked to dataset field
        passwordParamFound = True
        Exit For
        End If
      End If
      Next
      If not passwordParamFound Then
      'Subscription has Password field set to "Leave blank" or "Enter value"
      Dim pwdParameterValue As ParameterValue = New ParameterValue
      pwdParameterValue.Label = "PASSWORD"
      pwdParameterValue.Name = "PASSWORD"
      pwdParameterValue.Value = "PASSWORD"
      newParameterValues(extensionSettings.ParameterValues.Length) = pwdParameterValue		
      End If

      extensionSettings.ParameterValues = newParameterValues					
    End If
      
    Console.Write(description)
    If Not dataRetrievalPlan Is Nothing And TypeOf dataRetrievalPlan.Item Is DataSourceReference Then
      Dim item As DataSourceReference = dataRetrievalPlan.Item
      Dim ref As String = GetSnkPathRef(item.Reference)
      item.Reference = ref
    End If				
    RsSnk.CreateDataDrivenSubscription(snkReport.Path, extensionSettings, dataRetrievalPlan, description, eventType, GetMatchData(matchData, RsSnk), parameters)
    WriteSuccess()	

    If TypeOf dataRetrievalPlan.Item Is DataSourceDefinition Then
      Dim dsd As DataSourceDefinition = dataRetrievalPlan.Item
      If dsd.CredentialRetrieval = CredentialRetrievalEnum.Store Then
      Console.ForegroundColor = ConsoleColor.Yellow
            Dim warningMsg as String = "You need to re-enter the password for user '" + dsd.UserName + "' in data-driven subscription '" + description + "' contained in report '" + snkReport.Path + "'"
          Console.WriteLine(warningMsg)
          Console.ResetColor()

            LogErrorToFile("SUBSCRIPTIONS", "Warning: " + snkReport.Path, warningMsg)

      End If
    End If		
    
    Else
    Console.Write("Migrating subscription ")
    RsSrc.GetSubscriptionProperties(srcSub.SubscriptionID, extensionSettings, description, active, status, eventType, matchData, parameters)
    If extensionSettings.Extension = "Report Server FileShare" Then
      Dim newParameterValues() As [ParameterValueOrFieldReference] = New [ParameterValueOrFieldReference](extensionSettings.ParameterValues.Length) {}
      Array.Copy(extensionSettings.ParameterValues, newParameterValues, extensionSettings.ParameterValues.Length)
      'Standard subscriptions will always require adding PASSWORD parameter during migration
      Dim pwdParameterValue As ParameterValue = New ParameterValue
      pwdParameterValue.Label = "PASSWORD"
      pwdParameterValue.Name = "PASSWORD"
      pwdParameterValue.Value = "PASSWORD"
      newParameterValues(extensionSettings.ParameterValues.Length) = pwdParameterValue
      extensionSettings.ParameterValues = newParameterValues
    End If
    Console.Write(description)				
    RsSnk.CreateSubscription(snkReport.Path, extensionSettings, description, eventType, GetMatchData(matchData, RsSnk), parameters)
    WriteSuccess()		
    End If
    
    If extensionSettings.Extension = "Report Server FileShare" Then
    Dim username As String = "unknown"
    For Each pvofr As ParameterValueOrFieldReference In extensionSettings.ParameterValues
      If TypeOf pvofr Is ParameterValue Then
      Dim param As ParameterValue = pvofr
        If param.Name = "USERNAME" Then
        username = param.Value
        End If
      End If
    Next
    Console.ForegroundColor = ConsoleColor.Yellow
        Dim warningMsg as String = "You need to re-enter the password for user '" + username + "' in file share subscription '" + description + "' contained in report '" + snkReport.Path + "'."
      Console.WriteLine(warningMsg)
      Console.ResetColor()

        LogErrorToFile("SUBSCRIPTIONS", "Warning: " + snkReport.Path, warningMsg)

    End If
  Next

  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("SUBSCRIPTIONS", snkReport.Name, ex.Message)
  End Try
End Sub

'Migrate report's cache refresh plans
Sub MigrateCacheRefreshPlans(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Migrating cache refresh plans for report " + snkReport.Name + ": ")

  Try
  Dim srcPlans As CacheRefreshPlan() = RsSrc.ListCacheRefreshPlans(srcReport.Path)

  If srcPlans Is Nothing OrElse srcPlans.Length = 0 Then
    Console.WriteLine("0 items found.")
    Return
  End If
  Console.WriteLine(srcPlans.Length.ToString + " items found.")
  For Each srcPlan As CacheRefreshPlan In srcPlans

    Dim lastRunStatus As String = Nothing
    Dim state As CacheRefreshPlanState = Nothing
    Dim eventType As String = Nothing
    Dim matchData As String = Nothing
    Dim parameters As ParameterValue() = Nothing

    Try
    RsSrc.GetCacheRefreshPlanProperties(srcPlan.CacheRefreshPlanID, lastRunStatus, state, eventType, matchData, parameters)
    Console.Write("Migrating cache refresh plan " + srcPlan.Description)
    RsSnk.CreateCacheRefreshPlan(snkReport.Path, srcPlan.Description, eventType, GetMatchData(matchData, RsSnk), parameters)
    WriteSuccess()
    Catch e As Exception
    Console.ForegroundColor = ConsoleColor.Red
    Console.WriteLine(" ... FAILURE")
    Console.WriteLine(e.Message + Environment.NewLine)
    Console.ResetColor()
        LogErrorToFile("CACHEREFRESH", snkReport.Name, e.Message)
    End Try
  Next

  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("CACHEREFRESH", "Exception", ex.Message)
  End Try
End Sub

'Migrate report's parameter settings
Sub MigrateReportParameters(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Migrating parameters for report " + snkReport.Name)

  Try
  Dim srcParameters As ItemParameter() = RsSrc.GetItemParameters(srcReport.Path, Nothing, False, Nothing, Nothing)

  If srcParameters Is Nothing OrElse srcParameters.Length = 0 Then
    Console.WriteLine(" 0 items found.")
    Return
  End If

  RsSnk.SetItemParameters(snkReport.Path, srcParameters)
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("REPORTPARAMETERS", snkReport.Name, ex.Message)
  End Try
End Sub

'Migrate report's execution options
Sub MigrateReportExecutionOptions(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Migrating processing options for report " + snkReport.Name)

  Try
  Dim schedule As ScheduleDefinitionOrReference = Nothing
  RsSrc.GetExecutionOptions(srcReport.Path, schedule)

  If schedule Is Nothing Then
    Console.WriteLine(" ... 0 items found.")
    Return
  End If

  RsSnk.SetExecutionOptions(snkReport.Path, "Snapshot", schedule)
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("EXECUTIONOPTIONS", snkReport.Name, ex.Message)
  End Try
End Sub

'Migrate report's cache options
Sub MigrateReportCacheOptions(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Migrating cache refresh options for report " + snkReport.Name)

  Try
  Dim expirationDef As ExpirationDefinition = Nothing
  RsSrc.GetCacheOptions(srcReport.Path, expirationDef)

  If expirationDef Is Nothing Then
    Console.WriteLine(" ... 0 items found.")
    Return
  End If

  RsSnk.SetCacheOptions(snkReport.Path, True, expirationDef)
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("CACHEOPTIONS", snkReport.Name, ex.Message)
  End Try
End Sub

'Migrate report's history settings
Sub MigrateReportHistorySettings(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Migrating history settings for report " + snkReport.Name)

  Try
  Dim historySetting As ScheduleDefinitionOrReference = Nothing
  Dim keepExecutionSnapshots As Boolean = False
  RsSrc.GetItemHistoryOptions(srcReport.Path, keepExecutionSnapshots, historySetting)

  If historySetting Is Nothing Then
    Console.WriteLine(" ... 0 items found.")
    Return
  End If

  RsSnk.SetItemHistoryOptions(snkReport.Path, True, keepExecutionSnapshots, historySetting)

  Dim isSystem As Boolean = True
  Dim systemLimit As Integer = 0
  RsSrc.GetItemHistoryLimit(srcReport.Path, isSystem, systemLimit)
  RsSnk.SetItemHistoryLimit(snkReport.Path, isSystem, systemLimit)
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("HISTORYSETTINGS", snkReport.Name, ex.Message)
  End Try
End Sub


'Migrate report's policies
Sub MigrateItemPolicies(snkReport As CatalogItem, srcReport As CatalogItem)
  Console.Write("Migrating policies for item " + snkReport.Name)

  Try
  Dim inheritParent As Boolean = False
  Dim policies As Policy() = RsSrc.GetPolicies(srcReport.Path, inheritParent)

  If Not inheritParent And policies Is Nothing Then
    Console.WriteLine(" ...0 items found.")
    Return
  End If

  If Not inheritParent Then
    Dim mappedPolicies As Policy() = MapPolicies(policies)
    If mappedPolicies.Length > 0 Then
    RsSnk.SetPolicies(snkReport.Path, mappedPolicies)
    End If
  End If
    
  WriteSuccess()
  Catch ex As Exception
  Console.ForegroundColor = ConsoleColor.Red
  Console.WriteLine(" ... FAILURE:")
  Console.WriteLine(ex.Message + Environment.NewLine)
  Console.ResetColor()
    LogErrorToFile("ITEMPOLICIES", snkReport.Name, ex.Message)
  End Try
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''' Helper Functions '''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Helper function to link snk reference with relative paths
Function GetSnkPathRef(srcPath As String) As String
  Dim snkPath = srcPath

  If Not SrcIsNative Then
    'SharePoint integrated mode, will assume /sites server relative URL delimeter
    If srcSiteUrl isNot Nothing
      Dim delimeter = "/sites/"
    snkPath = srcPath.Remove(0, srcSiteUrl.IndexOf(delimeter) + delimeter.Length-1)
    End If
  Else
    'Treat path reference like other scenarios if not intergated mode
    snkPath = GetSnkPath(srcPath)
  End If

  Return snkPath
End Function
  
'Helper function to construct correctly formatted path
Function GetSnkPath(srcPath As String) As String
  If Not SrcIsNative Then
  srcPath = srcPath.Replace(SrcProtocol + "://" + SrcServer, "")
  End If

  Dim snkPath As String = SnkFolder

  If Not srcPath = "" Then
  If SrcFolder = SnkFolder Then ' exact relative folder paths assumed between source and target
    snkPath = srcPath
  Else
    snkPath = snkPath + IIf(SrcFolder = "/", srcPath, srcPath.Remove(0, SrcFolder.Length))
  End If		
  End If

  If snkPath.StartsWith("//") Then
  snkPath = snkPath.Substring(1)
  End If

  If Not SnkIsNative Then
  snkPath = SnkProtocol + "://" + SnkServer + snkPath
  End If

  Return snkPath
End Function

'Helper function to construct correctly formatted filename
Function GetSnkFilename(srcFilename As String, type As String) As String
  Dim currentFileExtension As String = GetFileNameExtension(srcFilename)

  If SnkIsNative And (GetExtension(type) = currentFileExtension) and currentFileExtension <> "" Then 'Cut off file extension, if present
  srcFilename = RemoveFileExtension(srcFilename)

  ElseIf Not SnkIsNative And currentFileExtension = "" Then 'Add file extension, if not present
  srcFilename = srcFilename + GetExtension(type)
  End If

  Return srcFilename
End Function

'Helper
Function GetFileNameExtension(filename As String) As String
  Dim index as Integer = -1
  index = filename.LastIndexOf(".")
  If index > 0
  Return filename.Substring(index)
  Else
  Return ""
  End If
End Function

'Helper
Function RemoveFileExtension(filename As String) As String
  Dim index As Integer = filename.LastIndexOf(".")
  Return filename.Substring(0, index)
End Function

'Helper function to construct correctly formatted path
Function GetSrcFolderPath() As String
  Return IIf(SrcIsNative, SrcFolder, SrcProtocol + "://" + SrcServer + SrcFolder)
End Function

'Returns the file extension for a given type
Function GetExtension(type As String) As String
  Dim ext As String = ""
  If type = "Report" Then
  ext = ".rdl"
  ElseIf type = "DataSource" Then
  ext = ".rsds"
  ElseIf type = "Component" Then
  ext = ".rsc"
  ElseIf type = "DataSet" Then
  ext = ".rsd"
  ElseIf type = "Model" Then
  ext = ".smdl"
  End If
  Return ext
End Function

Function GetMatchData(matchData As String, RsSnk As ReportingService2010) As String
  Dim name As String = Nothing
  For Each s As Schedule In schedules
  If s.ScheduleID = matchData Then
    name = s.Name
  End If
  Next
  If name = Nothing Then
  Return matchData
  End If

  Dim snkSchedules As Schedule() = RsSnk.ListSchedules(snkSiteUrl)
  For Each s As Schedule In snkSchedules
  If s.Name = name Then
    Return s.ScheduleID
  End If
  Next
  
  Return matchData
End Function

Function MapPolicies(policies As Policy()) As Policy()
  If(SrcIsNative = SnkIsNative) Then
  Return policies
  End If
  
  Dim mappedPolicies As System.Collections.Generic.List(Of Policy) = New System.Collections.Generic.List(Of Policy)
  
  For i As Integer = 0 To policies.Length - 1
  Dim mappedPolicy As Policy = new Policy
  mappedPolicy.GroupUserName = policies(i).GroupUserName
  mappedPolicy.Roles = MapRoles(policies(i).Roles)
  If mappedPolicy.Roles.Length > 0 Then
    mappedPolicies.Add(mappedPolicy)
  End If
  Next

  Return mappedPolicies.ToArray
End Function

Function MapRoles(roles As Role()) As Role()
  
  Dim mappedRoles As System.Collections.Generic.List(Of Role) = New System.Collections.Generic.List(Of Role)
  
  For i As Integer = 0 To roles.Length - 1
  Dim mappedRole As Role = new Role
  mappedRole.Name = MapRole(roles(i).Name)
  
  Try
    Dim taskIds As String() = RsSnk.GetRoleProperties(mappedRole.Name, snkSiteUrl, mappedRole.Description)
    mappedRoles.Add(mappedRole)
  Catch e as Exception
    Console.ForegroundColor = ConsoleColor.Yellow
    Console.WriteLine("Dropping '" + mappedRole.Name + ", which doesn't exist on the target server")
    Console.ResetColor()			
  End Try		
  Next	
  
  Return mappedRoles.ToArray
End Function

Function MapRole(roleName As String) As String	'These mapping rules may be adjusted to fit specific scenarios
  If SrcIsNative = SnkIsNative Then
  Return roleName
  Else If SrcIsNative Then 'Native to SP
  Select roleName 	
    Case "Content Manager"
    Return "Owners"
    Case "Publisher"
    Return "Members"
    Case "Browser"
    Return "Visitors"
  End Select
  Else If SnkIsNative Then 'SP to Native
  Select roleName 	
    Case "Owners"
    Return "Content Manager"
    Case "Members"
    Return "Publisher"
    Case "Visitors"
    Return "Browser"
  End Select
  End If	
  Return roleName
End Function

Function IIf(cond As Boolean, val1 As String, val2 As String) As String
  If cond Then
  Return val1
  Else
  Return val2
  End If
End Function

Sub WriteSuccess()
  Console.ForegroundColor = ConsoleColor.Green
  Console.WriteLine(" ... SUCCESS")
  Console.ResetColor()
End Sub

Function SortParentFoldersFirst(first As CatalogItem, second As CatalogItem) As Integer
  If first.TypeName = "Folder" And second.TypeName <> "Folder" Then
  Return -1
  ElseIf first.TypeName <> "Folder" And second.TypeName = "Folder" Then
  Return 1
  ElseIf first.TypeName <> "Folder" And second.TypeName <> "Folder" Then
  Return 0
  Else
  Return String.Compare(first.Path, second.Path)
  End If
End Function

Sub LogErrorToFile(source As String, errItem as String, message As String) 
  Dim writer as StreamWriter

  Try 
    Dim fileExists As Boolean = Not File.Exists(logFilePath) OrElse New FileInfo(logFilePath).Length = 0

    ' Create a StreamWriter object
    writer = New StreamWriter(logFilePath, True)

    If fileExists Then
      writer.WriteLine("DateTime, Source, ErrorItem, ErrorMessage")
    End If

    ' Format log message with timestamp
    Dim timeStamp As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    Dim formattedMessage as String = String.Format("{0}, {1}, {2}, ""{3}""", timeStamp, source, errItem, message.Replace(Environment.NewLine, " "))

    ' Write the log entry to file
    writer.WriteLine(formattedMessage)
  Catch ex As Exception
    Console.WriteLine(ex.Message)
  Finally
    If Not writer Is Nothing Then
      writer.Close()
    End If
  End Try
End Sub
