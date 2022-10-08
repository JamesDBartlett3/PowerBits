<#
    .SYNOPSIS
        Converts a PBIT file into Power BI Dataflow JSON import file.
    .DESCRIPTION
        It parses Power Query queries, their names, Power Query Editor groups, 
        and some additional properties from a PBIT file. Then it transforms 
        all the parsed information into a form which is used by Power BI Dataflows. 
        It is a JSON file used for import/export of dataflows.
    .INPUTS
        PBIT file
    .OUTPUTS
        JSON file with the name <PBIT original file name>.json
    .NOTES
        Version:        1.1
        Author:         Michal Dvorak (@nolockcz)
        Creation Date:  07.02.2020
        Purpose/Change: Fixed a bug with empty annotations
#>

param (
  [Parameter(Mandatory = $True)]
  [string]$inputFile
)

function HasProperty($object, $propertyName) {
  <#
      .SYNOPSIS
          Returns true, if an object contains a property
  #>

  return $propertyName -in $object.PSobject.Properties.Name
}

function ParseItems ($xmlBodyObject, $queries) {
  <#
      .SYNOPSIS
          Iterates over all non-empty items of xmlBodyObject (DataMashup file) and parse a query name, a query body, and some metadata.
  #>

  # iterate over all non-empty items
  foreach ($item in $xmlBodyObject.Items.Item | Where-Object { $_.StableEntries.InnerXml.Length -ne 0 }) {
      # get item path property, remove the prefix Section1/
      $itemPath = $item.ItemLocation.ItemPath -replace "Section1/", ""
      # if the item path is empty, skip this item
      if ($itemPath -eq "") { continue }
      # decode a query name from the item path
      $queryName = [uri]::UnescapeDataString($itemPath)
      
      # get QueryGroupID 
      # (the group hierarchy and group names are encrypted in the property QueryGroups, unfortunately)
      $queryGroupId = $item.StableEntries.Entry | Where-Object { $_.Type -eq "QueryGroupID" } | ForEach-Object { $_.Value }

      # get flag if a query is loaded or not
      $loadedToAnalysisServices = $item.StableEntries.Entry | Where-Object { $_.Type -eq "LoadedToAnalysisServices" } | ForEach-Object { $_.Value }
      $loadToReportDisabled = $item.StableEntries.Entry | Where-Object { $_.Type -eq "LoadToReportDisabled" } | ForEach-Object { $_.Value }
      $loadEnabled = (($loadedToAnalysisServices -eq "l1") -or ($loadToReportDisabled -ne "l1"))
      
      # parse a body of a query
      $queryBody = $null
      $lastAnalysisServicesFormulaText = $item.StableEntries.Entry | Where-Object { $_.Type -eq "LastAnalysisServicesFormulaText" } | ForEach-Object { $_.Value }        
      if ($lastAnalysisServicesFormulaText) {
          $formulaJsonObject = ConvertFrom-Json ($lastAnalysisServicesFormulaText.Substring(1))
          $queryBody = $formulaJsonObject.RootFormulaText
      }
      
      # create a query object - used later when creating a dataflow objects
      if (!(HasProperty $queries $queryName)) {
          $metadata = New-Object -TypeName psobject
          $metadata | Add-Member -MemberType NoteProperty -Name queryId -Value (New-Guid) # every query gets its GUID
          $metadata | Add-Member -MemberType NoteProperty -Name queryName -Value $queryName # query name
          $metadata | Add-Member -MemberType NoteProperty -Name loadEnabled -Value $loadEnabled # is load enabled?
          $metadata | Add-Member -MemberType NoteProperty -Name queryGroupId -Value $queryGroupId # group ID
          
          $queryObject = @{
              queryName = $queryName
              queryBody = $queryBody
              metadata = $metadata    
          }

          $queries | Add-Member -MemberType NoteProperty -Name $queryName -Value $queryObject
      }
      
      # if no query body, continue
      if (!$formulaJsonObject) { continue }
      # add query body to the query object
      $formulaJsonObject.ReferencedQueriesFormulaText.PSObject.Properties | ForEach-Object {
          if ((HasProperty $queries $_.Name) -and $_.Value) {
              $tmpQuery = $queries | Select-Object -ExpandProperty $_.Name
              $tmpQuery.queryBody = $_.Value
              $queries | Add-Member -MemberType NoteProperty -Name $_.Name -Value $tmpQuery -Force
          }
      }                
  }
}

function ParseQueriesFromPbit ($inputFile) {
  <#
      .SYNOPSIS
          Parse M queries from a PBIT file.
  #>

  # get file name without extension
  $zipDirectory = (Get-Item $inputFile).BaseName
  # create a new file name of a ZIP file
  $zipFileName = $zipDirectory + ".zip"

  try {
      # create a copy of a PBIT file with the extesion .ZIP
      Copy-Item -Path $inputFile -Destination $zipFileName

      # unzip
      Expand-Archive -Path $zipFileName -DestinationPath $zipDirectory
      # read the content of the file DataMashup in the unzipped file
      $DataMashupFileContent = Get-Content -Path "$zipDirectory\DataMashup" -Encoding UTF8 -Raw
      # read the content of the file DataModelSchema in the unzipped file
      $DataModelSchemaFileContent = Get-Content -Path "$zipDirectory\DataModelSchema" -Encoding Unicode -Raw
  }
  finally {
      # clean up all the temp files I have created
      Remove-Item -Path $zipFileName -Force
      Remove-Item -Path $zipDirectory -Force -Recurse
  }

  # get the innerXML string of <Items>
  $xmlBody = $DataMashupFileContent | Select-String -Pattern "<Items>.*</Items>" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }
  # parse the xml text to an XML object
  $xmlBodyObject = New-Object System.XML.XMLDocument
  $xmlBodyObject.LoadXml($xmlBody)    

  # initialize a new object containing all queries (properties are key-value pairs)
  $queries = New-Object -TypeName PSObject

  # two calls of ParseItems func because if a query contains only a reference to another query, the original query is empty
  # but I need also the body of the original query
  ParseItems $xmlBodyObject $queries
  ParseItems $xmlBodyObject $queries   

  # parse data model schema to get tables and columns which load is activated
  $DataModelSchemaJsonObject = ConvertFrom-Json $DataModelSchemaFileContent

  # for each table
  foreach ($table in $DataModelSchemaJsonObject.model.tables) {     
      # if there aren't any columns, continue
      if (!($table.columns)) { continue } 
      
      # get all columns of a table
      $columns = @()
      foreach ($column in $table.columns) {
          if ($column.type -eq "rowNumber") { continue }

          $columnObject = New-Object -TypeName psobject
          $columnObject | Add-Member -MemberType NoteProperty -Name name -Value $column.name
          $columnObject | Add-Member -MemberType NoteProperty -Name dataType -Value $column.dataType

          $columns += $columnObject
      }
      
      # add columns (which are loaded into the model) to queries
      if (!(HasProperty $queries $table.name)) { continue }
      $tmpQuery = $queries | Select-Object -ExpandProperty $table.name
      $tmpQuery | Add-Member -MemberType NoteProperty -Name columns -Value $columns
      $queries | Add-Member -MemberType NoteProperty -Name $table.name -Value $tmpQuery -Force
  }

  # select queries with non empty M codes
  $queriesAsArray = $queries.PSObject.Properties | ForEach-Object { $_.Value } | Where-Object { $null -ne $_.queryBody }
  # if there is a query containing the record #shared, show a warning
  foreach ($queryWithShared in $queriesAsArray | Where-Object { $_.queryBody -like "*#shared*" }) {
      Write-Warning ("Query {0} contains the record #shared and won't be migrated." -f $queryWithShared.queryName)
  }

  # ignore all queries containing the record #shared
  $queriesAsArray = $queriesAsArray# | Where-Object { $_.queryBody -notlike "*#shared*" }
  
  return $queriesAsArray
}

function ReplaceComments($str) {
  <#
      .SYNOPSIS
          Replace multiline comments like /* foo \r\n bar */ with //foo \r\n //bar
      .DESCRIPTION
          The PowerBI Dataflows don't support multiline comments and I need to replace them with one-line comments.
  #>
  $str = $str -replace "\r\n", "\r\n"
  $matchingStrings = $str | Select-String -Pattern "\/\*(.*?)\*\/" -AllMatches | ForEach-Object { $_.matches } | ForEach-Object { $_.value }

  foreach ($oldValue in $matchingStrings) {
      $newValue = $oldValue.Replace("\r\n", "\r\n//")
      $str = $str.Replace($oldValue, $newValue)
  }

  # replace the start and the end of a multiline comment
  $str = $str.Replace("/*", "//").Replace("*/", "")  
  
  return $str
}

function GenerateDocumentJson($queries) {
  <#
      .SYNOPSIS
          Joins M code of all queries into one long string.
  #>
  $documentJson = "section Section1;\r\n"
  $queries | ForEach-Object {
      $replacedComments = ReplaceComments($_.queryBody)
      $documentJson = $documentJson + 'shared #"' + $_.queryName + '" = ' + $replacedComments + ';\r\n'
  }
  return $documentJson
}

function GenerateMetadataObject($queries) {
  <#
      .SYNOPSIS
          Combines metadata of all queries and creates the property queriesMetadata.
  #>
  $metadataObject = New-Object -TypeName psobject    
  $queries | ForEach-Object {
      if (!$_.queryName) { continue }

      $metadataObject | Add-Member -MemberType NoteProperty -Name $_.queryName -Value $_.metadata
  }
  return $metadataObject
}

function GenerateAnnotations($queries) {
  <#
      .SYNOPSIS
          Generate annotations which are known as groups in Power Query Editor.
      .NOTES
          The file DataMashup contains a property called QueryGroups which contains group names and group hierarchy.
          Unfortunately, this property is somehow encrypted and I can't read the hierarchy and names.
          Therefore, I use groupIDs as names and no parentIDs. Your task is to rename the groups in Power BI Dataflow UI back to the original names.
  #>
  $values = @()
  $queries | Where-Object { $null -ne $_.metadata.queryGroupId } | ForEach-Object {
      $valueAsObject = @{
          id          = $_.metadata.queryGroupId
          name        = $_.metadata.queryGroupId # the name is unknown because it is encrypted, I use just the group ID instead of it
          description = $null
          parentId    = $null # because the hierarchy is encrypted in a PBIT file
          order       = 0
      }
      $values += ConvertTo-Json $valueAsObject -Compress
  }

  $values = $values | Sort-Object | Get-Unique
  
  $annotationObject = @{
      name  = "pbi:QueryGroups"
      value = ("[" + ($values -join ",") + "]")
  }
  
  return $annotationObject
}

function GeneratePbiMashup($queries) {
  <#
      .SYNOPSIS
          Generates the property PbiMashup.
  #>
  $pbiMashup = @{
      fastCombine        = $false
      allowNativeQueries = $false
      queriesMetadata    = (GenerateMetadataObject($queries))
      document           = (GenerateDocumentJson($queries))
  }

  return $pbiMashup
} 

function GenerateEntities($queries) {
  <#
      .SYNOPSIS
          Generates the property entities.
  #>
  $entities = @()
  foreach ($query in $queries) {        
      
      # if a query doesn't contain any columns, it isn't loaded into the model and can be ignored
      if (!$query.columns) { continue }

      $pbiRefreshPolicy = New-Object -TypeName psobject
      $pbiRefreshPolicy | Add-Member -MemberType NoteProperty -Name "`$type" -Value "FullRefreshPolicy"
      $pbiRefreshPolicy | Add-Member -MemberType NoteProperty -Name location -Value ([uri]::EscapeDataString($query.queryName + ".csv"))

      $entity = New-Object -TypeName psobject
      $entity | Add-Member -MemberType NoteProperty -Name "`$type" -Value "LocalEntity"
      $entity | Add-Member -MemberType NoteProperty -Name name -Value $query.queryName
      $entity | Add-Member -MemberType NoteProperty -Name description -Value ""
      $entity | Add-Member -MemberType NoteProperty -Name attributes -Value $query.columns
      $entity | Add-Member -MemberType NoteProperty -Name "pbi:refreshPolicy" -Value $pbiRefreshPolicy
      
      $entities += $entity 
  }

  return $entities
}

function GenerateMigrationString($inputFile) {
  <#
      .SYNOPSIS
          Combines all parts of the migration into an object. This object is then serialized to a JSON string.
  #>
  $queries = ParseQueriesFromPbit($inputFile)    

  $migrationObject = New-Object -TypeName psobject
  $migrationObject | Add-Member -MemberType NoteProperty -Name "name"  -Value $inputFile
  $migrationObject | Add-Member -MemberType NoteProperty -Name "description" -Value ""
  $migrationObject | Add-Member -MemberType NoteProperty -Name "version" -Value "1.0"
  $migrationObject | Add-Member -MemberType NoteProperty -Name "culture" -Value "en-US"
  $migrationObject | Add-Member -MemberType NoteProperty -Name "modifiedTime" -Value (Get-Date -Format o)
  $migrationObject | Add-Member -MemberType NoteProperty -Name "pbi:mashup" -Value (GeneratePbiMashup($queries))
  $migrationObject | Add-Member -MemberType NoteProperty -Name "entities" -Value @(GenerateEntities($queries))
  
  $annotations = @(GenerateAnnotations($queries))
  if ($annotations.Count -ne 0) {
      $migrationObject | Add-Member -MemberType NoteProperty -Name "annotations" -Value @($annotations)
  }

  $migrationString = (ConvertTo-Json $migrationObject -Depth 5).ToString().Replace("\\", "\")

  return $migrationString
}

# name of the output JSON file
$jsonOutputFileName = $inputFile.Replace(".pbit", ".json")

# generate the migration string from a PBIT file
GenerateMigrationString($inputFile) | Out-File $jsonOutputFileName -Encoding utf8