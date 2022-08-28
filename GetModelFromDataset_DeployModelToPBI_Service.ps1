$System_DefaultWorkingDirectory = "C:\Agent\_work\r3\a"
$tabular_editor_root_path = $System_DefaultWorkingDirectory
$client_id = ""
$client_secret = ""
$tenant_id = ""



# Download destination (root of PowerShell script execution path):
$DownloadDestination = join-path (get-location) "TabularEditor.zip"
$TabularEditorUrl = "https://github.com/TabularEditor/TabularEditor/releases/download/2.16.7/TabularEditor.Portable.zip"
# Download from GitHub:
Invoke-WebRequest -Uri $TabularEditorUrl -OutFile $DownloadDestination

# Unzip Tabular Editor portable, and then delete the zip file:
Expand-Archive -Path $DownloadDestination -DestinationPath (get-location).Path
Remove-Item $DownloadDestination


# we need to set Serialization Options to allow export to Folder via TE2
$serialization_options = '{
          "IgnoreInferredObjects": true,
          "IgnoreInferredProperties": true,
          "IgnoreTimestamps": true,
          "SplitMultilineStrings": true,
          "PrefixFilenames": false,
          "LocalTranslations": false,
          "LocalPerspectives": false,
          "LocalRelationships": false,
          "Levels": [
              "Data Sources",
              "Perspectives",
              "Relationships",
              "Roles",
              "Tables",
              "Tables/Columns",
              "Tables/Measures",
              "Translations"
          ]
      }'
$serialization_options | Out-File (Join-Path $tabular_editor_root_path "TabularEditor_SerializeOptions.json")
"Model.SetAnnotation(""TabularEditor_SerializeOptions"", ReadFile(@""$(Join-Path $tabular_editor_root_path "TabularEditor_SerializeOptions.json")""));" `
| Out-File (Join-Path $tabular_editor_root_path "ApplySerializeOptionsAnnotation.csx")


$FilePattern = Get-ChildItem (Join-Path $System_DefaultWorkingDirectory "_Power BI Dev Ops Pipeline-CI-Dataset\drop\MainDatasetDWH.pbix")
$filePath = $FilePattern.FullName
$workspaceName = "SEKO 360 [Development]_Test"
$DatasetName = "MainDatasetDWH"
$Switcher = "true"
$ServerName = "ig-ss-prod-we.database.windows.net"
#$DataBaseName = "ig-ss-prod-dwh"
$DataBaseName = "dwh-dev"


"var _ind=Model.Expressions[""Switcher""].Expression.IndexOf(""meta"");  
Model.Expressions[""Switcher""].Expression =""$Switcher""+Model.Expressions[""Switcher""].Expression.Substring(_ind-1);

var _indDB=Model.Expressions[""DataBaseName""].Expression.IndexOf(""meta"");  
Model.Expressions[""DataBaseName""].Expression =""\""$DataBaseName\""""+Model.Expressions[""DataBaseName""].Expression.Substring(_indDB-1);

var _indSN=Model.Expressions[""ServerName""].Expression.IndexOf(""meta"");  
Model.Expressions[""ServerName""].Expression =""\""$ServerName\""""+Model.Expressions[""ServerName""].Expression.Substring(_indSN-1);

Model.AllMeasures.FormatDax();
foreach(var m in Model.AllMeasures)
{
    m.Description =m.Expression;
}" | Out-File (Join-Path $tabular_editor_root_path "ChangeParam.csx")

"Model.SetAnnotation(""TabularEditor_SerializeOptions"", ReadFile(@""$(Join-Path $tabular_editor_root_path "TabularEditor_SerializeOptions.json")""));" `
| Out-File (Join-Path $tabular_editor_root_path "ApplySerializeOptionsAnnotation.csx")







[securestring]$sec_client_secret = ConvertTo-SecureString $client_secret -AsPlainText -Force
[pscredential]$credential = New-Object System.Management.Automation.PSCredential ($client_id, $sec_client_secret)
Connect-PowerBIServiceAccount -Credential $credential -ServicePrincipal -TenantId $tenant_id


Write-Host "Trying to publish file: $filePath 't to $workspaceName"
$temp_name = "$($FilePattern.BaseName)-$(Get-Date -Format 'yyyyMMddTHHmmss')"

$workspace = Get-PowerBIWorkspace -Name $workspaceName -ErrorAction SilentlyContinue
Write-Host "Uploading $($FilePattern.FullName) to $($workspace.Name)/$temp_name ... "
$report = New-PowerBIReport -Path $FilePattern.FullName -Name $temp_name -WorkspaceId $workspace.Id

Write-Host "Done!"

Write-Host "Getting PowerBI dataset ..."
$dataset = Get-PowerBIDataset -WorkspaceId $workspace.Id | Where-Object { $_.Name -eq $temp_name }
$connection_string = "powerbi://api.powerbi.com/v1.0/myorg/$($workspace.Name);initial catalog=$($dataset.Name)"

$login_info = "User ID=app:$client_id@$tenant_id;Password=$client_secret"

$ModelPath = Join-Path $FilePattern.DirectoryName "Model.bim"

$ParamsBIM = @(
    """Provider=MSOLAP;Data Source=$connection_string;$login_info"" ""$($dataset.Name)"""
    "-SCRIPT ""$(Join-Path $tabular_editor_root_path 'ApplySerializeOptionsAnnotation.csx')"" ""$(Join-Path $tabular_editor_root_path 'ChangeParam.csx')"""
    "-BIM ""$ModelPath"""
)


$executable = Join-Path $tabular_editor_root_path TabularEditor.exe
Write-Host "Starting Extraction PBIX metadata ..."
Write-Debug "$executable $params"
$p = Start-Process -FilePath $executable -Wait -NoNewWindow -PassThru -RedirectStandardOutput "$temp_name.log" -ArgumentList $ParamsBIM

if ($p.ExitCode -ne 0) {
    Write-Host "Failed to extract PBIX metadata from $($connection_string)!"
}
else {
    Write-Host "Extracted PBIX metadata to BIM $($ModelPath)!"
}




$DatasetTarget = Get-PowerBIDataset -WorkspaceId $workspace.Id | Where-Object { $_.Name -eq $DatasetName }

Write-Host "Overwriting 'name' and 'id' properties now ..."
$bim_json = Get-Content $ModelPath | ConvertFrom-Json
$bim_json.name = $DatasetTarget.Name
$bim_json.id = $DatasetTarget.ID.Guid
$bim_json | ConvertTo-Json -Depth 50 | Out-File $ModelPath
Write-Host "Overwrited 'name' and 'id' properties!"
          

if ($dataset -ne $null) {
    Write-Host "Removing temporary PowerBI dataset ..."
    Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($dataset.Id)" -Method Delete
}


$connection_string_del = "powerbi://api.powerbi.com/v1.0/myorg/$($workspaceName);Initial Catalog=$($DatasetName)"

$ParamsDeploy = @(
    """$ModelPath"""
    "-DEPLOY ""Data Source=$connection_string_del;$login_info"" ""$($DatasetName)"" -O -C -P -R -M -E -V"
)


$apiHeaders = Get-PowerBIAccessToken
$uri = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($DatasetTarget.ID.Guid)/refreshes?"+"`$top=1"
$result=Invoke-RestMethod -Uri $uri –Method GET -Headers $apiHeaders –Verbose
#checking if the dataset is refreshing?
if ($result.value.status -ne "Unknown") {

    Write-Host "Starting Deployment BIM to $($workspaceName) ... "

    $d = Start-Process -FilePath $executable -Wait -NoNewWindow -PassThru -RedirectStandardOutput "$temp_name.log" -ArgumentList $ParamsDeploy            

    if ($d.ExitCode -ne 0) {
        Write-Host "Failed to deploy BIM metadata to $($connection_string_del)!"
    }
    else {
        Write-Host "Deployed BIM metadata to $($workspaceName)!"
    }
}
else {
    Write-Host "Dataset is refreshing.It is forbidden to update model during refreshing!"
}