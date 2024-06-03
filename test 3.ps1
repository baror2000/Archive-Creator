#### Project For Beni

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
$baseDirectory = "C:\Project fot Beni\Tests\test 3"

# create folders
foreach ($FolderName in $FolderNames) {
    $folderPath = Join-Path -Path $baseDirectory -ChildPath $folderName

     if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory
        Write-Host "Folder '$folderName' created successfully."
    } else {
        Write-Host "Folder '$folderName' already exists."
    }
}


# create sub-folders

# download docs

# create word files