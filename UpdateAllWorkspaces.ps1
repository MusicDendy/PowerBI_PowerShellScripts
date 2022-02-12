#fill your local folder here:
$Path = "...\Bitbucket\powerbi\Reports\"

#$Workspaces=@("[Client]","[DC Operational]")
$Workspaces=@("Test Workspace")


#$ExludedFolders=@("Obsolete","Bespoke")
$ExludedFolders=@("Obsolete")

#if you want to update only reports that are in the TargetWorkspace and don't add new reports -> $AddNewReportsFlag=$false
$AddNewReportsFlag=$false
#if you want to add new reports in workspace -> $AddNewReportsFlag=$true

Foreach($Workspace in $Workspaces){

.\PublishReportToWorkspace.ps1 -Path $Path -ExludedFolders $ExludedFolders -TargetWorkspaceName $Workspace -AddNewReportsinTargetWorkspace $AddNewReportsFlag

}
