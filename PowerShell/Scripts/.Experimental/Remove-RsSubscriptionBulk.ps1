<#
.SYNOPSIS
Removes all subscriptions from a provided folder
 
.DESCRIPTION
Removes all subscriptions in a provided folder with an optional Recurse flag to include subscriptions in subfolders of the target folder.

.PARAMETER RsFolder
Target folder. This should be preceded by a /. Eg. '/Sales Reports'. Default is '/' (the root folder).
 
.PARAMETER Recurse
Flag to determine if subfolders should be included in the deletion. Leave blank to only remove subscriptions in the provided folder.
 
.EXAMPLE
Remove-RSSubscriptionBulk -RSfolder '/' -Recurse
 
This will remove all subscriptions in an entire instance
 
.EXAMPLE
Remove-RSSubscriptionBulk -RSfolder '/Sales Reports' -Confirm
 
This will remove all subscriptions in the Sales Reports folder only. It will not affect subfolders. It will also prompt before each subscription deletion.
 
.NOTES
General notes
#>
 
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Parameter(Mandatory)][string]$ServerName,
  [Parameter()][int]$PortNumber = 443,
  [Alias('ItemPath', 'Path')][Parameter()][string]$RsFolder = '/',
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
    Write-Verbose "$($subs.Count) Subscriptions will be deleted."
    foreach ($sub in $subs) {
      if ($pscmdlet.ShouldProcess($sub.Path, "Delete Subscription with ID: $($sub.SubscriptionID)")) {
        $proxy.DeleteSubscription($sub.SubscriptionID)
      }                
      Write-Verbose "Subscription Deleted: $($sub.SubscriptionID)"
    }
  } catch {
    throw (New-Object System.Exception("Failed to delete items in '$RsFolder': $($_.Exception.Message)", $_.Exception))
  }  
}    
end {
  if ($subs) {
    Write-Host "Deleted $($subs.Count) Subscriptions"
  }
}