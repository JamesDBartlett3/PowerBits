#---------------------------------------------
# Author: Craig Porteous @cporteous
# Modified by: James Bartlett @jamesdbartlett3
# Synopsis:	Remove a specific user/group from 
#		    all SSRS folders and reports. 
#		    Excludes inherited folders.
#---------------------------------------------
 
$ReportServerName = ''
$ReportServerUri = "https://$ReportServerName/ReportServer/ReportService2010.asmx?wsdl"
$InheritParent = $true
$GroupUserName = ''
[array]$itemTypes = ("Folder", "Report", "DataSet", "DataSource", "Model", "Resource", "LinkedReport", "ReportPart", "PowerBIReport")

$rsProxy = New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential

#List out all subfolders under the parent directory
$Items = $rsProxy.ListChildren("/", $true) |
         Where-Object {$_.typeName -in $itemTypes} |
         Select-Object Path | Sort-Object -Property Path

#Iterate through every folder
ForEach($Item in $Items) {

  Write-Output "Revoking $GroupUserName's access to $($Item.Path)..."

  # WHY DOES THIS WORK?!
  # It shouldn't be changing anything, but it apparently does.
  # Is the [ref]$InheritParent parameter doing something weird?
 	$Policies = $rsProxy.GetPolicies($Item.Path, [ref]$InheritParent)
 
 	#Skip over folders marked to Inherit permissions. No changes needed.
 	if($InheritParent -eq $false) {

 		#List out ALL policies on folder but do not include the policy for the specified user/group
 		$Policies = $Policies | Where-Object { $_.GroupUserName -ne $GroupUserName }
 
 		#Set the folder's policies to this new set of policies
 		$rsProxy.SetPolicies($Item.Path, $Policies)
    }

 }