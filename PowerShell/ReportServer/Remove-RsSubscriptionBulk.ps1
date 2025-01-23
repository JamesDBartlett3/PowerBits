<#
.SYNOPSIS
Removes all ReportServer subscriptions in a specified folder
 
.DESCRIPTION
This removes all ReportServer subscriptions in a specified folder, with a Recurse flag to include all subfolders, and an option to either Delete or Disable the subscriptions.
 
.PARAMETER RsFolder
Target folder. This should be preceded by a /. Eg. '/Sales Reports'. It is possible to set the Root folder using '/'

.PARAMETER ServerName
Name of the server on which the reportserver is running (e.g. reportserver.example.com)

.PARAMETER PortNumber
Port number on which the reportserver is running (default: 443)

.PARAMETER ReportServerInstanceName
Name of the reportserver instance (default: ReportServer)

.PARAMETER Action
Delete or Disable the subscriptions. Default is Disable. If Delete is specified, the subscriptions will be permanently removed.
 
.PARAMETER Recurse
Flag to determine if all subfolders should be included or only the target folder
 
.EXAMPLE
# This will permanently delete all subscriptions in the enteire ReportServer instance on 'reportserver.example.com'
Remove-RSSubscriptionBulk -RSfolder '/' -ServerName 'reportserver.example.com' -Action 'Delete' -Recurse
 
.EXAMPLE
# This will disable all subscriptions in the 'Sales Reports' folder and all subfolders on 'reportserver.example.com'
Remove-RSSubscriptionBulk -RSfolder '/Sales Reports' -ServerName 'reportserver.example.com' -Action 'Disable' -Recurse
 
.EXAMPLE
# This will disable all subscriptions in the 'Sales Reports' folder only. It will not affect subfolders.
Remove-RSSubscriptionBulk -RSfolder '/Sales Reports' -ServerName 'reportserver.example.com' -Action 'Disable'
 
.NOTES
THIS SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND. USE AT YOUR OWN RISK.
#>
 
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Alias('ItemPath', 'Path')][Parameter(Mandatory)][string]$RsFolder,
  [Parameter(Mandatory)][string]$ServerName,
  [Parameter()][int]$PortNumber = 443,
  [Parameter()][string]$ReportServerInstanceName = "ReportServer",
  [Parameter()][ValidateSet("Delete", "Disable")][string]$Action = "Disable",
  [Parameter()][switch]$Recurse
)
begin {
  #Requires -Modules ReportingServicesTools
  $Proxy = (New-RsWebServiceProxy -ReportServerUri "https://$($ServerName):$($PortNumber)/$($ReportServerInstanceName)/ReportService2010.asmx?wsdl")
}
process {
  try {
    if($Recurse){
      Write-Verbose "Recurse flag set. Return all subscriptions in Folder:$($RsFolder) and sub-folders."
      $subs = $Proxy.ListSubscriptions($RsFolder)
    }
    else{
      Write-Verbose "Recurse flag not set. Return all subscriptions in Folder:$($RsFolder) only"
      $subs = $Proxy.ListSubscriptions($RSFolder) | Where-Object {$_.Path -eq "$($RsFolder)/$($_.Report)"}
    }
  }
  catch {
    throw (New-Object System.Exception("Failed to retrieve items in '$RsFolder': $($_.Exception.Message)", $_.Exception))
  }
  try {
    # Declare action strings for verbose output
    $actionPresentTense = $Action.Substring(0, $Action.Length - 1) + "ing"
    $actionLowerCase = $Action.ToLower()
    $actionLowerCasePastTense = $actionLowerCase + "d"
    # Execute the specified action on each subscription
    Write-Verbose "$($subs.Count) Subscriptions will be $actionLowerCasePastTense."
    foreach($sub in $subs){
      if ($pscmdlet.ShouldProcess($sub.Path, "$actionPresentTense Subscription with ID: $($sub.SubscriptionID)")) {
        if ($Action -eq "Disable") {
          $Proxy.DisableSubscription($sub.SubscriptionID)
        } elseif ($Action -eq "Delete") {
          $Proxy.DeleteSubscription($sub.SubscriptionID)
        } else {
          throw (New-Object System.Exception("Invalid action specified: $Action"))
        }
      }
      Write-Verbose "Subscription $($actionLowerCasePastTense): $($sub.SubscriptionID)"
    }
  }
  catch {
    throw (New-Object System.Exception("Failed to $actionLowerCase items in '$RsFolder': $($_.Exception.Message)", $_.Exception))
  }
}
end {
  Write-Verbose "Completed."
}