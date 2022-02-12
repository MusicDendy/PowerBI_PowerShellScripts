
$ExludedWorkspaces=@("Perfomance Test","SEKO 360 [Admin Usage]","Test Workspace")


$LocalPath="C:\Reports\file.csv"

$table=@()

 #Login-PowerBI
Connect-PowerBIServiceAccount | Out-Null
$workspaces=Get-PowerBIWorkspace | Where-Object {$_.Name -notin $ExludedWorkspaces}

 Foreach ( $workspace in $workspaces){
  
  $reports=Get-PowerBIReport -WorkspaceId $workspace.id.Guid
    foreach ($report in $reports){
        $row="" | select Workspace,Report
        $row.workspace=$workspace.Name
        $row.report=$report.Name
        $table+=$row;

    }
 }
 $table | Out-File -FilePath $LocalPath

