### Project for Beni By Baro

### This script creates folders, sownloads files and creates word summery files according to an Excel file
### please note that the Excel file should be called data.xlsx

# Install required modules
Install-Module ImportExcel
Install-Module -Name PSWriteWord -RequiredVersion 1.0.1

# Import the modules
Import-Module ImportExcel
Import-Module PSWriteWord

# Read Excel data
$currentPath = get-location
$excelData = Import-Excel -Path "$currentPath\Data.xlsx"

# Initialize counter for progress tracking
$totalItems = $excelData.Count
$currentItem = 0

# Download function
function Download-File {
    param([string]$Url, [string]$FilePath)
    Invoke-WebRequest -Uri $Url -OutFile $FilePath
}

# Create summary document function
function Create-Summary-Document {
    param([string]$Content, [string]$OutputPath)
    try {
        $WordDocument = New-WordDocument $OutputPath
        Add-WordText -WordDocument $WordDocument -Text $Content -
        Save-WordDocument $WordDocument
        Write-Host "Created Word document: $OutputPath"
    } catch {
        Handle-Error "Failed to create document at $OutputPath"
    }
}


# Create summary document function
function Create-Summary-Document {
    param([string]$Content, [string]$OutputPath)
    try {
        $WordDocument = New-WordDocument $OutputPath
        Add-WordText -WordDocument $WordDocument -Text $Content
        Save-WordDocument $WordDocument
        Write-Host "Created Word document: $OutputPath"
    } catch {
        Handle-Error "Failed to create document at $OutputPath"
    }
}


# Main Script
foreach ($row in $excelData) {

 # Progress Bar
    $currentItem++
    Write-Progress -Activity "Downloading files" -Status "Downloading files for $($row.ID)" -PercentComplete (($currentItem / $totalItems) * 100)

# Create a folder for each ID
    $folderPath = "$currentPath\folders\$($row.ID)"
    New-Item -ItemType Directory -Path $folderPath
    
    # Create subfolders "ק-0 ק-1 ק-2 חוק" within the folder
    foreach ($subFolder in ("ק-0", "ק-1", "ק-2", "נוסח חוק")) {
        $subFolderPath = Join-Path -Path $folderPath -ChildPath $subFolder
        New-Item -ItemType Directory -Path $subFolderPath
    }

    # Download files from URLs
    
    foreach ($url_ק0 in $row.'נוסח ק0') {
        Download-File $url_ק0 "$folderPath\ק-0\נוסח ק0.docx"
        }
    foreach ($url_ק1 in $row.'נוסח ק1') {
        Download-File $url_ק1 "$folderPath\ק-1\נוסח ק1.pdf"
        }
    foreach ($url_ק2 in $row.'נוסח ק2') {
        Download-File $url_ק2 "$folderPath\ק-2\נוסח ק2.pdf"
        }
    foreach ($url_חוק in $row.'נוסח חוק') {
        Download-File $url_חוק "$folderPath\נוסח חוק\נוסח חוק.docx"
        }

    #creating Summery files

    foreach ($sum_ק0 in $row.'תקציר ק0') {
        Create-Summary-Document $sum_ק0 "$folderPath\ק-0\תקציר ק0.docx"
        }

     foreach ($sum_ק1 in $row.'תקציר ק1') {
        Create-Summary-Document $sum_ק1 "$folderPath\ק-1\תקציר ק1.docx"
        }

     foreach ($sum_ק2 in $row.'תקציר ק2') {
        Create-Summary-Document $sum_ק2 "$folderPath\ק-2\תקציר ק2.docx"
        }

     foreach ($sum_חוק in $row.'תקציר חוק') {
        Create-Summary-Document $sum_חוק "$folderPath\נוסח חוק\תקציר חוק.docx"
        }
}

Write-Progress -Activity "Download Complete" -Status "All files downloaded successfully" -Completed
    