
$TargetWorkspaceName="Test Workspace"

#Login-PowerBI
Connect-PowerBIServiceAccount | Out-Null

$target_workspace = Get-PowerBIWorkspace -Name $TargetWorkspaceName -ErrorAction SilentlyContinue

$newReports = Get-PowerBIReport -WorkspaceId $target_workspace.Id

Write-Host "Starting Remove Reports from $($TargetWorkspaceName)" -ForegroundColor Red;
Write-Host "----------------";

Foreach ($newReport in $newReports) {

        Remove-PowerBIReport -Id $newReport.Id -WorkspaceId $target_workspace.Id 

        Write-Host "Removed $($newReport.Name)"

}
Write-Host "----------------";
Write-Host $newReports.Count " Reports Removed Succesfully" -ForegroundColor Green;
Write-Host "";
