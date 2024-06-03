### Project for Beni By Baro

# Install required modules
Install-Module ImportExcel
Install-Module -Name PSWriteWord -RequiredVersion 1.0.1

# Import the modules
Import-Module ImportExcel
Import-Module PSWriteWord

# Read Excel data
$excelData = Import-Excel -Path "C:\Project for Beni\Data.xlsx"

# Initialize counter for progress tracking
$totalItems = $excelData.Count
$currentItem = 0

# Loop through each row in the Excel file
foreach ($row in $excelData) {

    # Increment progress counter
    $currentItem++

    # Create a folder for each ID
    $folderPath = "C:\Project for Beni\Tests\test 7\folders\$($row.ID)"
    New-Item -ItemType Directory -Path $folderPath
    
    # Create subfolders "ק-0 ק-1 ק-2 חוק" within the folder
    foreach ($subFolder in ("ק-0", "ק-1", "ק-2", "נוסח חוק")) {
        $subFolderPath = Join-Path -Path $folderPath -ChildPath $subFolder
        New-Item -ItemType Directory -Path $subFolderPath
    }

    # Download files from URLs
    #Write-Progress -Activity "Downloading files" -Status "Downloading files for $($row.ID)" -PercentComplete (($currentItem / $totalItems) * 100)

    #foreach ($url_ק0 in $row.'נוסח ק0') {
    #    Invoke-WebRequest -Uri $url_ק0 -OutFile "$folderPath\ק-0\נוסח ק0.docx"
    #    }
    #foreach ($url_ק1 in $row.'נוסח ק1') {
    #    Invoke-WebRequest -Uri $url_ק1 -OutFile "$folderPath\ק-1\נוסח ק1.pdf"
    #    }
    #foreach ($url_ק2 in $row.'נוסח ק2') {
    #    Invoke-WebRequest -Uri $url_ק2 -OutFile "$folderPath\ק-2\נוסח ק2.pdf"
    #    }
    #foreach ($url_חוק in $row.'נוסח חוק') {
    #    Invoke-WebRequest -Uri $url_חוק -OutFile "$folderPath\נוסח חוק\נוסח חוק.docx"
    #    }

    #creating Summery files
    foreach ($sum_ק0 in $row.'תקציר ק0') {

        $outputPath = "$folderPath\ק-0\תקציר ק0.docx"
        $WordDocument = New-WordDocument $outputPath
        $content = $sum_ק0

        Add-WordText -WordDocument $WordDocument -Text $content
        Save-WordDocument $WordDocument
        Write-Host "Created Word document: $outputPath"
        }


     foreach ($sum_ק1 in $row.'תקציר ק1') {
        $outputPath = "$folderPath\ק-1\תקציר ק1.docx"
        $WordDocument = New-WordDocument $outputPath
        $content = $sum_ק1

        Add-WordText -WordDocument $WordDocument -Text $content
        Save-WordDocument $WordDocument
        Write-Host "Created Word document: $outputPath"
        }

     foreach ($sum_ק2 in $row.'תקציר ק2') {
        $outputPath = "$folderPath\ק-2\תקציר ק2.docx"
        $WordDocument = New-WordDocument $outputPath
        $content = $sum_ק2

        Add-WordText -WordDocument $WordDocument -Text $content
        Save-WordDocument $WordDocument
        Write-Host "Created Word document: $outputPath"
        }

     foreach ($sum_חוק in $row.'תקציר חוק') {
        $outputPath = "$folderPath\תקציר חוק\תקציר חוק.docx"
        $WordDocument = New-WordDocument $outputPath
        $content = $sum_חוק

        Add-WordText -WordDocument $WordDocument -Text $content
        Save-WordDocument $WordDocument
        Write-Host "Created Word document: $outputPath"
        }
}

Write-Progress -Activity "Download Complete" -Status "All files downloaded successfully" -Completed
    