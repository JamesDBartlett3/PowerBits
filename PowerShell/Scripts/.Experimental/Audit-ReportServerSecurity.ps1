#---------------------------------------------
# Author: 	James Bartlett @jamesdbartlett3@techhub.social
# Original author: 	Craig Porteous @cporteous
# Synopsis: List out all SSRS (native mode)
#		    folders and their security policies,
#		    then output dataset to Excel or CSV file
# TODO: The inheritance logic currently only checks if the GroupUserName matches. 
#				Need to also check if the roles match.
# TODO: Check for redundant individual level permissions
# 			(i.e. if a user has the same permissions as a group they are a member of)
# TODO: Refactor with a recursive function to handle nested folders
# TODO: Add parameter to turn Active Directory feature on/off
#---------------------------------------------

[CmdletBinding()]
param (
	[Parameter(Mandatory)][string]$ReportServerName,
	[Parameter(Mandatory)][string]$ReportServerPort,
	[Parameter()][string]$OutputDirectory = $PSScriptRoot,
	[Parameter()][string]$OutputFileNamePrefix = "ReportServer_SecurityAudit",
	[Parameter()][ValidateSet("Excel","CSV")][string]$OutputFileFormat = "Excel",
	[Parameter()][boolean]$InheritParent = $true,
	[Parameter()][string]$SSRSroot = "/"
)

$CurrentPSVersion = [Version]::new($PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor)
$MaxPSVersion = [Version]::new(5, 1)
$IsVerbose = $PSCmdlet.MyInvocation.BoundParameters["Verbose"]

if ($CurrentPSVersion -gt $MaxPSVersion) {
	Write-Host -ForegroundColor Red "This script requires PowerShell version 5.1 or lower. Exiting script..."
	Exit
}

$Excel = ($OutputFileFormat -eq "Excel")

if ($Excel) {
	try {
		Import-Module -Name ImportExcel | Out-Null
	} catch {
		Write-Host -ForegroundColor Yellow "OutputFileFormat parameter set to 'Excel' but ImportExcel module not found. Install now? (Y/N)"
		$choice = Read-Host
		if($choice.ToUpper() -eq "Y") {
			Install-Module -Name ImportExcel -Scope CurrentUser
		} else {
			Write-Host -ForegroundColor Red "Cannot export to Excel without ImportExcel module. Exiting script..."
			Exit
		}
	}
}

$currentDate = Get-Date -UFormat "%Y-%m-%d"
$OutputFilePath = Join-Path -Path $OutputDirectory -ChildPath ($OutputFileNamePrefix + "__(" + $ReportServerName + "_" + $currentDate + ")")
$ReportServerUri = "https://" + $ReportServerName + ":" + $ReportServerPort + "/ReportServer/ReportService2010.asmx?wsdl"
$rsPerms = @()
$rsResult = @()
$Separator = ""

# Create a new Web Service Proxy object to connect to SSRS
$rsProxy = New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential

# List out all subfolders under the parent directory and Select their "Path"
$folderList = $rsProxy.ListChildren($SSRSroot, $InheritParent) | Where-Object {$_.TypeName -EQ "Folder"}

# Iterate through every folder 
foreach($folder in $folderList) {

	# Return all policies on this folder
	$Policies = $rsProxy.GetPolicies($folder.Path, [ref]$InheritParent)

	# Compare the policies to the parent folder and see if they are inherited or not
	$parentFolderPath = (Split-Path -Path $folder.Path -Parent).Replace("\", "/")
	$ParentPolicies = $rsProxy.GetPolicies($parentFolderPath, [ref]$InheritParent)
	$Policies = $Policies | Select-Object *,@{Name="Inherited";Expression={"False"}}
	$Policies | Where-Object {$ParentPolicies.GroupUserName -contains $_.GroupUserName} | ForEach-Object {$_.Inherited = "True"}

	# For each policy, add details to an array
	if ($Policies.Count -gt 0) {
		if ($IsVerbose) {
			$Separator = "-" * (($folder.Path).length + 8)
			Write-Host $Separator -ForegroundColor Blue
			Write-Host "Folder:" $folder.Path
			Write-Host $Separator	-ForegroundColor Blue
			Write-Host "↳ Policies:"
		}
		foreach($rsPolicy in $Policies) {
			$groupUserName = $rsPolicy.GroupUserName.Split("\")[-1];
			$roles = $rsPolicy.Roles | Select-Object -Property Name
			if ($IsVerbose) {
				Write-Host "  ↳" $rsPolicy.GroupUserName
				foreach($role in $roles) {
					Write-Host "    ▸" $role.Name}
			}
			$roleString = $roles.Name -join "|"
			[array]$rsResult = New-Object PSObject -Property @{
				"ID" = $folder.ID;
				"Path" = $folder.Path;
				# Remove the domain name from the GroupUserName value
				"GroupUserName" = $groupUserName;
				"Roles" = $roleString;
				"Inherited" = $rsPolicy.Inherited
			}
			$rsPerms += $rsResult
		}
		if ($IsVerbose) {Write-Host "$Separator`n`n" -ForegroundColor Blue}
	}
}

# Add file extension to output file path
$OutputFilePath += if($Excel) {".xlsx"} else {".csv"}

if ($IsVerbose) {
	$Separator = "=" * (($OutputFilePath.length) + 14)
	Write-Host $Separator -ForegroundColor Green
	Write-Host "Writing file: $OutputFilePath"
	Write-Host $Separator -ForegroundColor Green
}

if ($IsVerbose) {
	Write-Host "Checking for disabled accounts in Active Directory..." -ForegroundColor Yellow
}

# Loop through all unique GroupUserName values in the $rsPerms array, and check if it is active in Active Directory
foreach($rsPerm in $rsPerms | Select-Object -Property GroupUserName -Unique) {
	$ADGroup = Get-ADGroup -Filter "$("GroupCategory -eq 'Security' -and Name -eq ' " + $rsPerm.GroupUserName + "'")"
	if (-not $ADGroup) {
		$ADUser = Get-ADUser -Filter "$("SamAccountName -eq ' " + $rsPerm.GroupUserName + "'")" -Properties Enabled
		if (-not $ADUser.Enabled) {
			# If the user is disabled, add a tag to the GroupUserName value in the $rsPerms array
			$rsPerms | Where-Object {$_.GroupUserName -eq $rsPerm.GroupUserName} | ForEach-Object {$_.GroupUserName += " (Disabled)"}
		}
	}
}

# Create a new array with the results
$result = $rsPerms | Select-Object -Property ID, Path, GroupUserName, Roles, Inherited

# Output array to file
if($Excel) {
	$result | Export-Excel -TableStyle Medium1 -FreezeTopRow -AutoFilter -AutoSize -Show -Path $OutputFilePath
} else {
	$result | Export-Csv -Path $OutputFilePath -NoTypeInformation
	Invoke-Item $OutputFilePath
}