<#
-------------------------------------------------------
Script Name: 			SSRS Grant Access by Item.ps1
Original author:	Craig Porteous @cporteous
Modified by: 			James Bartlett @jamesdbartlett3
Synopsis:					Grant access to a user or group on 
                  a specific item in SSRS
-------------------------------------------------------
#>

param(
		[Parameter(Mandatory=$true)][string]$ReportServerName,
		[Parameter(Mandatory=$true)][string]$ItemPath,
		[Parameter(Mandatory=$true)][string]$GroupUserName,
		[Parameter(Mandatory=$false)]
			[ValidateSet('Browser', 'Content Manager', 'My Reports', 'Publisher', 'Report Builder', 'System Administrator', 'System User')]
			[string]$RoleName = 'Browser'
)

# If running as PowerShell Core (version 6+), re-launch in a new Windows PowerShell (version 5.1) window
if ($PSVersionTable.PSEdition -eq 'Core') {
	Start-Process powershell.exe -ArgumentList "-NoExit","-File `"$PSCommandPath`"" -Verb RunAs
	exit
}

# Types of items to which the change will be applied
[array]$itemTypes = ("Folder", "Report", "DataSet", "DataSource", "Model", "Resource", "LinkedReport", "ReportPart", "PowerBIReport")

$ReportServerUri = "https://$ReportServerName/ReportServer/ReportService2010.asmx?wsdl"
$InheritParent = $true
$rsProxy = New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential
$Type = $rsProxy.GetType().Namespace;
$policyType = "{0}.Policy" -f $Type;
$roleType = "{0}.Role" -f $Type;

# Set FolderPath to the parent folder of the ItemPath
$FolderPath = $ItemPath.Substring(0, $ItemPath.LastIndexOf("/"))

# List out all subfolders under the parent directory
$Items = $rsProxy.ListChildren($FolderPath, $True) | `
         Select-Object -Property TypeName, Path, ID, Name | `
         Where-Object -Property TypeName -in $itemTypes | `
         Where-Object -Property Path -eq $ItemPath

# Iterate through every item
foreach($Item in $Items) {

	$Policies = $rsProxy.GetPolicies($Item.Path, [ref]$InheritParent)

	# Skip over folders marked to Inherit permissions. No changes needed.
	if($InheritParent -eq $false) {

		# Return all policies that contain the user/group we want to add
		$Policy = $Policies | `
		    Where-Object { $_.GroupUserName -eq $GroupUserName } | `
		    Select-Object -First 1

		# Add a new policy if doesnt exist
		if (-not $Policy) {

		    $Policy = New-Object ($policyType)
		    $Policy.GroupUserName = $GroupUserName
		    $Policy.Roles = @()

			#Add new policy to the folder's policies
		    $Policies += $Policy
		}

		# Add the role to the new Policy
		$r = $Policy.Roles | `
            Where-Object { $_.Name -eq $RoleName } | `
	        Select-Object -First 1
	    if (-not $r) {

	        $r = New-Object ($roleType)
	        $r.Name = $RoleName
	        $Policy.Roles += $r
    	}

		# Apply policy to target items
		$rsProxy.SetPolicies($Item.Path, $Policies);

	}

}