<#
.SYNOPSIS
Removes, enables, or disables all subscriptions in a folder in Reporting Services

.DESCRIPTION
Performs the specified action (Remove, Enable, Disable) on all subscriptions in a provided folder with an optional Recurse flag to include subscriptions in subfolders of the target folder.

.PARAMETER ServerName (Mandatory)
The name of the Reporting Services server. This should be the server name or IP address. Do not include the protocol or port number.

.PARAMETER Action (Mandatory)
Specifies the action to perform on the subscriptions. Valid values are 'Enable', 'Disable', and 'Delete'.

.PARAMETER PortNumber (Optional)
Port number on which the report server is running. Default is 443.

.PARAMETER RsFolder (Optional)
Target folder in the report server. Should always start with a forward slash, e.g., '/Sales Reports'. Default is '/' (the root folder).

.PARAMETER Recurse (Optional)
Flag to determine if subfolders should be included in the action. Leave blank to only affect subscriptions in the provided folder.

.EXAMPLE
Modify-RSSubscriptionBulk -RSfolder '/' -Action 'Delete' -Recurse -Confirm

This will remove all subscriptions from all reports in the root folder and all subfolders, prompting the user to confirm before each subscription is deleted.

.EXAMPLE
Modify-RSSubscriptionBulk -RSfolder '/Sales Reports' -Action 'Disable'

This will disable all subscriptions on all reports in the '/Sales Reports' folder only. It will not affect subfolders.

.EXAMPLE
Modify-RSSubscriptionBulk -RSfolder '/Sales Reports' -Action 'Enable' -Recurse

This will enable all subscriptions on all reports in the '/Sales Reports' folder and all subfolders.

.NOTES
  ACKNOWLEDGEMENTS
    - Thanks to my wife (@likeawednesday@techhub.social) for her support and encouragement.
    - Thanks to the PowerShell and Power BI/Fabric communities for being so awesome.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Parameter(Mandatory)][string]$ServerName,
  [Parameter(Mandatory)][ValidateSet('Enable', 'Disable', 'Delete')][string]$Action,
  [Parameter()][int]$PortNumber = 443,
  [Alias('Path', 'Folder', 'TargetFolder')][Parameter()][string]$RsFolder = '/',
  [Parameter()][switch]$Recurse
)

begin {
  #Requires -Modules ReportingServicesTools
  $uri = "https://$($ServerName):$($PortNumber)/ReportServer/ReportService2010.asmx?wsdl"
  $proxy = (New-RsWebServiceProxy -ReportServerUri $uri)
}    
process {
  try {
    if ($Recurse) {
      Write-Verbose "Recurse flag set. Return all subscriptions in Folder:$($RsFolder) and sub-folders"
      $subs = $proxy.ListSubscriptions($RsFolder)
    } else {
      Write-Verbose "Recurse flag not set. Return all subscriptions in Folder:$($RsFolder) only"
      $subs = $proxy.ListSubscriptions($RSFolder) | Where-Object {$_.Path -eq "$($RsFolder)/$($_.Report)"}                        
    }
  } catch {
    throw (New-Object System.Exception("Failed to retrieve items in '$RsFolder': $($_.Exception.Message)", $_.Exception))
  }
  try {
    Write-Verbose "$($subs.Count) Subscriptions will be processed for action: $Action."
    foreach ($sub in $subs) {
      $methodName = "${Action}Subscription"
      if ($pscmdlet.ShouldProcess($sub.Path, "$Action Subscription '$($sub.Description)' (ID: $($sub.SubscriptionID))")) {
        $proxy.$methodName($sub.SubscriptionID)
      }
      Write-Verbose "Subscription $($Action): $($sub.SubscriptionID)"
    }
  } catch {
    throw (New-Object System.Exception("Failed to process items in '$RsFolder': $($_.Exception.Message)", $_.Exception))
  }  
}    
end {
  if (!$WhatIfPreference.IsPresent -and $subs) {
    Write-Host "$Action completed for $($subs.Count) Subscriptions"
  }
}