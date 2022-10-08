Get-ChildItem -Recurse "*.pbix" | 
  Select-Object -Property FullName, BaseName, @{
    l="PbixInBaseNameFolder"; e={$_.BaseName -eq (Split-Path (Split-Path $_.FullName -Parent) -Leaf)}
  } | Select-Object -Property PbixInBaseNameFolder, @{l="FullName"; e={$_.FullName}}, @{
    l="SourceDir"; e={
      if ($_.PbixInBaseNameFolder){
        Join-Path (Split-Path $_.FullName -Parent) "src"
      } else {
        Join-Path (Split-Path $_.FullName -Parent) (Join-Path $_.BaseName "src")
      }
    }
  } | Format-List