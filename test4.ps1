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
$baseDirectory = "C:\Project fot Beni\Tests\test 4\folders"

# Define the names of the subfolders to create in each new folder
$subfolders = @("ק-0", "ק-1", "ק-2", "חוק")

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
}


# download docs

# create word files