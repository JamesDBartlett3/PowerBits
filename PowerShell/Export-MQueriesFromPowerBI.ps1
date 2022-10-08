#Finding the portnumber on which the $Embedded$ tabular model is running on
$embedded = "$env:LOCALAPPDATA\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces"
$ports = Get-ChildItem $embedded -rec | Where-Object {$_.Name -eq "msmdsrv.port.txt"}
$port = Get-Content $ports.FullName -Encoding Unicode

#Getting the data sources from the $Embedded$ tabular model
# TODO: Loop thru ports until connection succeeds 
[xml]$db = Invoke-ASCmd -Server:localhost:$($port[1]) -Query:"SELECT * from `$SYSTEM.TMSCHEMA_DATA_SOURCES"
$db.return.root.row
$cs = $db.return.root.row.ConnectionString

#Cleaning up the connection string
$b64 = $cs.Split(";")[-1].Trim("Mashup=").Trim("""")
$bytes = [Convert]::FromBase64String($b64)
$temp = "c:\temp"

#Output to a binary ZIP file
[IO.File]::WriteAllBytes("$temp\a.zip", $bytes)

#Unzip
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("$temp\a.zip","$temp\a")
#TADA!
# ($m = Get-Content $temp\a\Formulas\Section1.m)