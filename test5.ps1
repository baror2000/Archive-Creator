#### Project For Beni by Baro

# Install the moduels
Install-Module -Name ImportExcel -Scope CurrentUser

# Import the modules
Import-Module ImportExcel

# Define the path to the Excel file and the sheet name
$excelFilePath = "C:\Project fot Beni\Data.xlsx"
$sheetName = "Sheet1" 

# Import the Excel data
$data = Import-Excel -Path $excelFilePath

# Get the list of folder names
$folderNames = $data."ID"

# Define the base directory where you want to create the folders
$baseDirectory = "C:\Project fot Beni\Tests\test 5\folders"

# Define the names of the subfolders to create in each new folder
$subfolders = @("ק-0", "ק-1", "ק-2", "חוק")

# Function to download a file from a URL
function Download-File {
    param (
        [string]$url,
        [string]$outputPath
    )
    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath
        Write-Host "Downloaded: $outputPath"
    } catch {
        Write-Host "Failed to download: $url"
    }
}


# create folders
foreach ($FolderName in $FolderNames) {
    $folderPath = Join-Path -Path $baseDirectory -ChildPath $folderName

     if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory
        Write-Host "Folder '$folderName' created successfully."
    } else {
        Write-Host "Folder '$folderName' already exists."
    }
    
    # create sub-folders
     foreach ($subfolder in $subfolders) {
            $subfolderPath = Join-Path -Path $folderPath -ChildPath $subfolder
            if (-not (Test-Path -Path $subfolderPath)) {
                New-Item -Path $subfolderPath -ItemType Directory
                Write-Host "Subfolder '$subfolder' created in folder '$folderName'."
            } else {
                Write-Host "Subfolder '$subfolder' already exists in folder '$folderName'."
            }
        }
    # Define mappings for URL and summary columns
            $mappings = @(
                @{ Subfolder = "ק-0"; UrlColumn = "C"; SummaryColumn = "B"; FileName = "K0.docx" },
                @{ Subfolder = "ק-1"; UrlColumn = "E"; SummaryColumn = "D"; FileName = "K1.pdf" },
                @{ Subfolder = "ק-2"; UrlColumn = "G"; SummaryColumn = "F"; FileName = "K2.pdf" },
                @{ Subfolder = "חוק"; UrlColumn = "I"; SummaryColumn = "H"; FileName = "Law.docx" }
            )
    # Process each mapping
            foreach ($mapping in $mappings) {
                $subfolderPath = Join-Path -Path $folderPath -ChildPath $mapping.Subfolder

                # Download the file from URL
                    foreach ($row in $data) {  # Iterate through each row in $data
                    $url = $row.$($mapping.UrlColumn)
                        if ($url -ne $null -and $url -ne "") {
                        $outputFilePath = Join-Path -Path $subfolderPath -ChildPath $mapping.FileName  # Fix file name
                        Download-File -url $url -outputPath $outputFilePath
                        }
                }
                # Save the summary content as a Word document
                # $summaryContent = $row."$($mapping.SummaryColumn)"
                #if ($summaryContent -ne $null -and $summaryContent -ne "") {
                #    $wordFilePath = Join-Path -Path $subfolderPath -ChildPath "$($mapping.FileName)_Summary.docx"
                #    Create-WordDocument -content $summaryContent -outputPath $wordFilePath
                #}
            }
}
