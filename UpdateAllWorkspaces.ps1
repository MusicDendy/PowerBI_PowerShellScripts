#fill your local folder here:
$Path = "C:\Users\d.shumilin\OneDrive - Itransition, CJSC\Documents\Project\Bitbucket\powerbi\Reports\"

#$Workspaces=@("SEKO 360 [Client]","SEKO 360 [DC Operational]","SEKO 360 [Indigina]","SEKO 360 [Bespoke]","SEKO 360 [Client & DC]")
$Workspaces=@("Test Workspace")


#$ExludedFolders=@("Obsolete","Bespoke")
$ExludedFolders=@("Obsolete")

#if you want to update only reports that are in the TargetWorkspace and don't add new reports -> $AddNewReportsFlag=$false
$AddNewReportsFlag=$false
#if you want to add new reports in workspace -> $AddNewReportsFlag=$true

Foreach($Workspace in $Workspaces){

.\PublishReportToWorkspace.ps1 -Path $Path -ExludedFolders $ExludedFolders -TargetWorkspaceName $Workspace -AddNewReportsinTargetWorkspace $AddNewReportsFlag

}