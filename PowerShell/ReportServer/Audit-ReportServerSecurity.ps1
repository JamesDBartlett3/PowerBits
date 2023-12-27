#----------------------------------------------------------------------------------------------------------------------------------
# Author: 	James Bartlett @jamesdbartlett3@techhub.social
# Original author: 	Craig Porteous @cporteous
# Synopsis: List out all SSRS (native mode) folders and their security policies, then output dataset to Excel or CSV file
#
# TODO FIRST: Refactor to use ReportingServicesTools module and add support for PowerShell Core
# ---------------------------------------------------------------------------------------------------------------------------------
# TODO: The inheritance logic currently only checks if the GroupUserName matches. Need to also check if the roles match.
# TODO: Add item-level permissions to the output
# TODO: Add activity check to see if the user/group has accessed the folder in the last X days
# TODO: Add activity check for activity on items by all users
# TODO: Add GroupUserType column (i.e. User, ADGroup, etc.)
# TODO: Check for redundant individual level permissions (i.e. if a user has the same permissions as a group they are a member of)
# TODO: Refactor with a recursive function to handle nested folders
# TODO: Add parameter to turn Active Directory feature on/off
# TODO: Passthru param for routing output to another script or function
#----------------------------------------------------------------------------------------------------------------------------------

[CmdletBinding()]
param (
  [Parameter(Mandatory)][string]$ReportServerName,
  [Parameter(Mandatory)][string]$ReportServerPort,
  [Parameter()][string]$OutputDirectory = $PSScriptRoot,
  [Parameter()][string]$OutputFileNamePrefix = "ReportServer_SecurityAudit",
  [Parameter()][ValidateSet("Excel","CSV")][string]$OutputFileFormat = "Excel",
  [Parameter()][boolean]$InheritParent = $true,
  [Parameter()][boolean]$IncludeADCheck = $true,
  [Parameter()][string]$SSRSroot = "/"
)

$CurrentPSVersion = [Version]::new($PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor)
$MaxPSVersion = [Version]::new(5, 1)
$IsVerbose = $PSCmdlet.MyInvocation.BoundParameters["Verbose"]

if ($CurrentPSVersion -gt $MaxPSVersion) {
  Write-Host -ForegroundColor Red "This script requires PowerShell version $MaxPSVersion or lower. Exiting script..."
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

if ($IncludeADCheck) {
  try {
    Import-Module -Name ActiveDirectory | Out-Null
  } catch {
    Write-Host -ForegroundColor Yellow "IncludeADCheck parameter set to 'True' but ActiveDirectory module not found. Install now? (Y/N)"
    $choice = Read-Host
    if($choice.ToUpper() -eq "Y") {
      Install-Module -Name ActiveDirectory -Scope CurrentUser
    } else {
      Write-Host -ForegroundColor Red "Cannot check Active Directory without ActiveDirectory module. Exiting script..."
      Exit
    }
  }
}

$currentDate = Get-Date -UFormat "%Y%m%d_%H%M%S"
$OutputFilePath = Join-Path -Path $OutputDirectory -ChildPath ($OutputFileNamePrefix + "_-_" + $ReportServerName + "_" + $currentDate)
$ReportServerUri = "https://" + $ReportServerName + ":" + $ReportServerPort + "/ReportServer/ReportService2010.asmx?wsdl"
$rsPerms = @()
$rsResult = @()
$Separator = ""

# Create a new Web Service Proxy object to connect to SSRS
$rsProxy = New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential

# List out all subfolders under the parent directory and Select their "Path"
$folderList = $rsProxy.ListChildren($SSRSroot, $true) | Where-Object {$_.TypeName -EQ "Folder"}

# Iterate through every folder 
foreach($folder in $folderList) {

  # Return all policies on this folder
  $Policies = $rsProxy.GetPolicies($folder.Path, [ref]$InheritParent)

  # Compare the policies to the parent folder and see if they are inherited or not
  $parentFolderPath = (Split-Path -Path $folder.Path -Parent).Replace("\", "/")
  $ParentPolicies = $rsProxy.GetPolicies($parentFolderPath, [ref]$InheritParent)
  $Policies = $Policies | Select-Object *,@{Name="Inherited"; Expression={$false}}
  $Policies | Where-Object {$ParentPolicies.GroupUserName -contains $_.GroupUserName} | ForEach-Object {$_.Inherited = $true}

  # For each policy, add details to an array
  if ($Policies.Count -gt 0) {
    if ($IsVerbose) {
      $Separator = "-" * (($folder.Path).length + 8)
      Write-Host $Separator -ForegroundColor Blue
      Write-Host "Folder:" $folder.Path
      Write-Host $Separator	-ForegroundColor Blue
      Write-Host " Policies:"
    }

    # For each policy, add details to an array
    foreach($rsPolicy in $Policies) {
      # Remove the domain name from the GroupUserName value
      $groupUserName = $rsPolicy.GroupUserName.Split("\")[-1];
      # Get the domain name from the GroupUserName value
      $groupUserDomain = $rsPolicy.GroupUserName.Split("\")[0];
      $roles = $rsPolicy.Roles | Select-Object -Property Name
      if ($IsVerbose) {
        Write-Host "  |-" $rsPolicy.GroupUserName
        foreach($role in $roles) {
          Write-Host "  |   |-" $role.Name}
      }
      $roleString = $roles.Name -join "|"
      [array]$rsResult = New-Object PSObject -Property @{
        "FolderID" = $folder.ID;
        "FolderPath" = $folder.Path;
        "GroupUserDomain" = $groupUserDomain;
        "GroupUserName" = $groupUserName;
        "Disabled" = $false;
        "Roles" = $roleString;
        "Inherited" = $rsPolicy.Inherited
      }
      $rsPerms += $rsResult
    }
    if ($IsVerbose) {Write-Host "$Separator`n`n" -ForegroundColor Blue}
  }
}

# TODO: wrap in ad check logic
if ($IsVerbose) {
  Write-Host "Checking for disabled accounts in Active Directory..." -ForegroundColor Yellow
}

# Loop through all unique GroupUserName values in the $rsPerms array, and check if it is active in Active Directory
foreach($rsPerm in $rsPerms | Where-Object GroupUserDomain -ne 'BUILTIN' | Select-Object -Property GroupUserName -Unique) {
  $ADGroup = Get-ADGroup -Filter "$("GroupCategory -eq 'Security' -and SamAccountName -eq ' " + $rsPerm.GroupUserName + "'")"
  if (-not $ADGroup) {
    $ADUser = Get-ADUser -Filter "$("SamAccountName -eq ' " + $rsPerm.GroupUserName + "'")" -Properties Enabled
    if (-not $ADUser.Enabled) {
      # If the user is disabled, mark it as such in the GroupUserDisabled property in $rsPerms array
      $rsPerms | Where-Object {$_.GroupUserName -eq $rsPerm.GroupUserName} | ForEach-Object {$_.Disabled = $true}
    }
  }
}

# Create a new array with the results
$result = $rsPerms | Select-Object -Property FolderID, FolderPath, GroupUserDomain, GroupUserName, Disabled, Roles, Inherited

# Add file extension to output file path
$OutputFilePath += if($Excel) {".xlsx"} else {".csv"}

if ($IsVerbose) {
  $Separator = "=" * (($OutputFilePath.length) + 14)
  Write-Host $Separator -ForegroundColor Green
  Write-Host "Writing file: $OutputFilePath"
  Write-Host $Separator -ForegroundColor Green
}

# Output array to file
if($Excel) {
  $result | Export-Excel -TableStyle Medium1 -FreezeTopRow -AutoFilter -AutoSize -Path $OutputFilePath
} else {
  $result | Export-Csv -Path $OutputFilePath -NoTypeInformation
}

# Open the file
Invoke-Item $OutputFilePath