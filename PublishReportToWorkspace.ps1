<#
.Synopsis
    Copies the contents of a Power BI workspace to another Power BI workspace.
.Description
    Copies the contents of a Power BI workspace to another Power BI workspace.
	This script creates the target workspace if it does not exist.
    This script uses the Power BI Management module for Windows PowerShell. If this module isn't installed, install it by using the command 'Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser'.
.Parameter SourceWorkspaceName
    The name of the workspace you'd like to copy the contents from.
.Parameter TargetWorkspaceName
    The name of the workspace you'd like to copy to. You must have edit access to the specified workspace.
.Parameter CreateTargetWorkspaceIfNotExists
    A flag to indicate if the script should create the target workspace if it doesn't exist. The default is to create the target workspace.
.Example
    PS C:\> .\PublishReportToWorkspace.ps1 -TargetWorkspaceName "Targetworkspace" -SourceWorkspaceName "My Workspace" 
	Copies the contents of the current user's personal workspace to a new workspace called "Copy of My Workspace".
#>

[CmdletBinding()]
param
(
    
[string] $Path = "C:\Bitbucket\powerbi\Reports\",
[object] $ExludedFolders=@("Obsolete"),
[string] $SourceWorkspaceName="SEKO 360 [Production]",
[string] $DatasetName="MainDatasetDWH",
[string] $DevelopmentWorspace="SEKO 360 [Development]",
[object] $ExcludedFiles = @("Theme.json", "Template.pbix", "Template.pbit", "RLS Validating.pbix"),
[bool]   $CreateTargetWorkspaceIfNotExists = $true,
#[bool]   $UsedDevelopDatasetinGit = $true,
[bool]   $AddNewReportsinTargetWorkspace=$true,


#Target Folder for sync reports
[string] $TargetWorkspaceName

)

#region Helper Functions 

function Assert-ModuleExists([string]$ModuleName) {
    $module = Get-Module $ModuleName -ListAvailable -ErrorAction SilentlyContinue
    if (!$module) {
        Write-Host "Installing module $ModuleName ..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
        Write-Host "Module installed"
    }
    elseif ($module.Version -ne '1.0.0' -and $module.Version -le '1.0.410') {
        Write-Host "Updating module $ModuleName ..."
        Update-Module -Name $ModuleName -Force -ErrorAction Stop
        Write-Host "Module updated"
    }
}

#endregion

# ==================================================================
# PART 1: Verify that the Power BI Management module is installed
#         and authenticate the current user.
# ==================================================================
Assert-ModuleExists -ModuleName "MicrosoftPowerBIMgmt"
#Login-PowerBI
Connect-PowerBIServiceAccount | Out-Null

# ==================================================================
# PART 2: Getting source and target workspace
# ==================================================================
# STEP 2.1: Get the source workspace
$source_workspace_ID = ""
while (!$source_workspace_ID) {
    $source_workspace_name = if (-not($SourceWorkspaceName)) {
        Read-Host -Prompt "Enter the name of the workspace you'd like to copy from" 
    }
    else {
        $SourceWorkspaceName 
    }

    if ($source_workspace_name -eq "My Workspace") {
        $source_workspace_ID = "me"
        break
    }

    $workspace = Get-PowerBIWorkspace -Name $source_workspace_name -ErrorAction SilentlyContinue

    if (!$workspace) {
        Write-Warning "Could not get a workspace with that name. Please try again, making sure to type the exact name of the workspace"  
    }
    else {
        $source_workspace_ID = $workspace.id
    }
}

# STEP 2.2: Get the target workspace
$target_workspace_ID = "" 
while (!$target_workspace_ID) {
    $target_workspace_name = if (-not($TargetWorkspaceName)) {
        Read-Host -Prompt "Enter the name of the workspace you'd like to copy to" 
    }
    else {
        $TargetWorkspaceName 
    }
	
    $target_workspace = Get-PowerBIWorkspace -Name $target_workspace_name -ErrorAction SilentlyContinue

    if (!$target_workspace -and $CreateTargetWorkspaceIfNotExists -eq $true) {
        $target_workspace = New-PowerBIWorkspace -Name $target_workspace_name -ErrorAction SilentlyContinue
    }

    if (!$target_workspace -or $target_workspace.isReadOnly -eq "True") {
        Write-Error "Invalid choice: you must have edit access to the workspace."
        break
    }
    else {
        $target_workspace_ID = $target_workspace.id
    }

    if (!$target_workspace_ID) {
        Write-Warning "Could not get a workspace with that name. Please try again with a different name."  
    } 
}

# ==================================================================
# PART 3: Copying reports and datasets 
# ==================================================================


# STEP 3.1: Get the reports from the target workspace
 $ReportsInTargetWorkspace = Get-PowerBIReport -WorkspaceId $target_workspace_ID | Where-Object {$_.Name -notin @("Usage Metrics","Report Usage Metrics Report")}


 #Write-Host "Getting Dataset Information from $DevelopmentWorspace";
 $dev_dataset = Get-PowerBIDataset -WorkspaceId (Get-PowerBIWorkspace -Name $DevelopmentWorspace).Id.Guid | Where-Object {$_.Name -in $DatasetName} -ErrorAction Stop;

# STEP 3.2: If we have reports in target workspace we have to rebind to Dev Dataset

if ($ReportsInTargetWorkspace){
    Write-Host "Rebinding to Dev" ;
    Write-Host "----------------";
    Foreach ($newReport in $ReportsInTargetWorkspace) {

            ## SEND REQUEST 
            $requestBody = @{datasetId = $dev_dataset.Id.Guid};
            $requestBodyJson = $requestBody | ConvertTo-Json -Compress;  
          
            $headers = Get-PowerBIAccessToken;
            $result = Invoke-RestMethod `
                -Headers $headers `
                -Method "Post" `
                -ContentType "application/json" `
                -Uri "https://api.powerbi.com/v1.0/myorg/groups/$($target_workspace.Id.Guid)/reports/$($newReport.Id.Guid)/Rebind" `
                -Body $requestBodyJson `
                -Timeout 3600 `
                -ErrorAction Stop;

            Write-Host "Rebinded $($newReport.Name)"

    }
    Write-Host "----------------";
    Write-Host $ReportsInTargetWorkspace.Count " Reports Rebinded Succesfully" -ForegroundColor Green;
    Write-Host "";
}

# STEP 3.3: Get the reports from the Bitbucket Folder

$ObsoletePath=$ExludedFolders | ForEach-Object {"$Path$_"}
$reports = Get-ChildItem -Path $Path -Recurse -Include "*.pbix" | Where-Object {$_.name -notin $ExcludedFiles} | Where-Object {$_.Directory -notin $ObsoletePath}

if (!$AddNewReportsinTargetWorkspace){

    #Update only reports that are in the TargetWorkspace and don't add new reports

    $reports=$reports | Where-Object {$_.BaseName -in $ReportsInTargetWorkspace.Name} 

}

Write-Host "Publishing to target workspace $TargetWorkspaceName"
Write-Host "----------------"; 

Foreach ($report in $reports) {
    $report_name = $report.BaseName
    $report_path = $report.FullName


try {
        Write-Host "Published $report_name" 

        $new_report = New-PowerBIReport -WorkspaceId $target_workspace_ID -Path $report_path -Name $report_name -ConflictAction CreateOrOverwrite

        #if ($new_report) {
            # keep track of the report and dataset IDs
        #    $report_id_mapping[$report_id] = $new_report.id
        #    $dataset_id_mapping[$dataset_id] = $new_report.datasetId
        #}
    }
    catch [Exception] {
        Write-Error "== Error: failed to import PBIX"

        $exception = Resolve-PowerBIError -Last
        Write-Error "Error Description:" $exception.Message
        continue
    }
}


Write-Host "----------------";
Write-Host $reports.Count" Reports Published" -ForegroundColor Green
Write-Host ""

# ==================================================================
# PART 4: Rebind new Reports To Source Dataset.
# ==================================================================

# STEP 4.1: Get the reports from the target workspace
 $ReportsInTargetWorkspace = Get-PowerBIReport -WorkspaceId $target_workspace_ID | Where-Object {$_.Name -notin @("Usage Metrics","Report Usage Metrics Report")}
 Write-Host "";
 Write-Host "Rebinding Dataset to $SourceWorkspaceName" ;
 Write-Host "----------------";
 $dataset =Get-PowerBIDataset -WorkspaceId $source_workspace_ID.Guid | Where-Object {$_.Name -in $DatasetName} -ErrorAction Stop;

Foreach ($newReport in $ReportsInTargetWorkspace) {

        ## SEND REQUEST 
        $requestBody = @{datasetId = $dataset.Id.Guid};
        $requestBodyJson = $requestBody | ConvertTo-Json -Compress;  
          
        $headers = Get-PowerBIAccessToken;
        $result = Invoke-RestMethod `
            -Headers $headers `
            -Method "Post" `
            -ContentType "application/json" `
            -Uri "https://api.powerbi.com/v1.0/myorg/groups/$($target_workspace.Id.Guid)/reports/$($newReport.Id.Guid)/Rebind" `
            -Body $requestBodyJson `
            -Timeout 3600 `
            -ErrorAction Stop;

        Write-Host "Rebinded $($newReport.Name)";

}
Write-Host "----------------";
Write-Host $ReportsInTargetWorkspace.Count " Reports Deployed to $TargetWorkspaceName and Rebinded Succesfully to $SourceWorkspaceName" -ForegroundColor Green;
Write-Host "----------------";