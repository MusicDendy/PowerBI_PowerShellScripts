#Parameters
$SourcePath = $env:Build_ArtifactStagingDirectory
$TargetPath = Join-Path (Split-Path -Parent $SourcePath) "Target"

$7zip = "C:\Program Files\7-Zip\7z.exe"

$pbix_files = Get-ChildItem -Path $SourcePath -Include *.pbix -Recurse

#Get Current Date to Calculate LastDay and FirstDay of Month
$CURRENTDATE = Get-Date -Format "yyyy-MM-ddT00:00:00"
$FIRSTDAY = Get-Date $CURRENTDATE -Day 1
$LASTDAY = Get-Date $FIRSTDAY.AddMonths(1)
$FIRSTDAYOFMONTH = Get-Date $FIRSTDAY -Format "yyyy-MM-ddT00:00:00"
$LASTDAYOFMONTH = Get-Date $LASTDAY -Format "yyyy-MM-ddT00:00:00"

#check log file exists,if log file exists then remove it
if ((Test-Path (Join-Path (Split-Path -Parent $SourcePath)  "temp.log"))) {
    Remove-Item  -Force -Path (Join-Path (Split-Path -Parent $SourcePath)  "temp.log") -Confirm:$false
}
#check Target folder exists,if Target folder exists then remove it and create new one
If ((Test-Path $TargetPath)) {
    Remove-Item   -Path $TargetPath -Force -Recurse -Confirm:$false
    New-Item -Path $TargetPath -ItemType Directory  -Force
}

foreach ($pbix_file in $pbix_files) {

    #Make temp folder
    $tempDir = $env:TEMP + "\PBI TEMP"
    $tempfolderfromfile = Join-Path $tempDir $pbix_file.BaseName

    #create new tempfolder in tempdir
    If (!(Test-Path $tempfolderfromfile))
    { New-Item -Path $tempfolderfromfile -ItemType Directory  -Force }

    #copy pbix to tempdir
    Copy-Item -Path $pbix_file.FullName -Destination $tempDir -Force

    $oldName = Get-ChildItem -Path (Join-Path $tempDir $pbix_file.Name)

    # Hard-coded file.zip as it doesn't matter what the filename is called
    $newName = Join-Path $tempDir ($pbix_file.BaseName + ".zip")

    # Replace .pbi[xt] with .zip
    # This is because Expand-Archive only works on .zip files
    if (Test-Path $newName) {
        Get-Item $newName  | Remove-Item  -Force
    }
    $oldName | Rename-Item -NewName ($pbix_file.BaseName + ".zip") -Force

    # Extract file.zip
    Expand-Archive $newName -DestinationPath  $tempfolderfromfile -Force

    #Get Content from Layout file in pbix
    $Layout = Get-Content -Raw -Path ($tempfolderfromfile + "\Report\Layout") -Encoding unicode

    #find all dates with pattern
    $datePattern = [Regex]::new('\d\d21-\d\d-\d\dT\d\d:\d\d:\d\d')
    $matchespatternBefore = $datePattern.Matches($Layout)
    $DTUnique = $matchespatternBefore.Value | Sort-Object -Unique

    #modify Layout file
    $i = 0
    foreach ($DTItem in $DTUnique) {
        #Logging before changing
        $pbix_file.FullName + "|" + $DTItem | Out-File -Append -FilePath (Join-Path (Split-Path -Parent $SourcePath)  "temp.log")
        #replace startDate in DateFilter
        if ($i -eq 0) {
            $Layout = $Layout.Replace($DTItem, $FIRSTDAYOFMONTH)
        }
        else {
            $Layout = $Layout.Replace($DTItem, $LASTDAYOFMONTH)
        }
        $i++
    }

    #checking after modification
    $pbix_file.FullName
    $datePattern = [Regex]::new('\d\d\d\d-\d\d-\d\dT\d\d:\d\d:\d\d')
    $matchespatternAfter = $datePattern.Matches($Layout)
    $matchespatternAfter.Value | Sort-Object -Unique

    #remove BOM encodint before saving.
    #Parameters specify whether to use the big endian byte order and byte order mark(BOM) UnicodeEncodint(Endianness,BOM)
    #Big endian byte order: 00 00 00 41,Little endian byte order: 41 00 00 00
    $Utf16NoBomEncoding = New-Object -TypeName System.Text.UnicodeEncoding($False, $False) #UnicodeEncodint(Little endian,No BOM)
    [System.IO.File]::WriteAllText(($tempfolderfromfile + "\Report\Layout"), $Layout, $Utf16NoBomEncoding)

    #remove protection in pbix file and save it.
    Get-Item ($tempfolderfromfile + "\SecurityBindings")  | Remove-Item  -Force

    $ContentTypesPath = $tempfolderfromfile + "\[Content_Types].xml"
    $ContentTypes = [System.IO.File]::ReadAllText($ContentTypesPath)
    $ContentTypes = $ContentTypes.Replace('<Override PartName="/SecurityBindings" ContentType="" />', '')

    $Utf8BomEncoding = New-Object System.Text.UTF8Encoding($True)  # Use BOM in file
    [System.IO.File]::WriteAllText($ContentTypesPath, $ContentTypes, $Utf8BomEncoding)

    #creating subdirectory in Target folder
    if (!(Test-Path -Path ($TargetPath + $pbix_file.DirectoryName.Replace($SourcePath, "")))) {
        New-Item -ItemType Directory -Path ($TargetPath + $pbix_file.DirectoryName.Replace($SourcePath, "")) -Force
    }

    #create pbix file from tempfolderfile
    & $7zip a -tzip -mx=6 ($TargetPath + $pbix_file.FullName.Replace($SourcePath, "")) ($tempfolderfromfile + "\*") -r
}

#Delete temp
Get-ChildItem -Path ($tempDir) -Recurse | Remove-Item -Force -Recurse -Confirm:$false

#replace source files from target folder
Copy-Item -Path ($TargetPath + "\*") -Destination $SourcePath -Force -Recurse
#remove target folder
Remove-Item -Path $TargetPath -Force -Recurse -Confirm:$false