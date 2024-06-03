### Excel Script for Beni

# Install the moduel
Install-Module -Name ImportExcel -Scope CurrentUser

# Define the Excel file path
$excelFile = "C:\Project fot Beni\data.xlsx"

# Import the sheet data
$data = Import-Excel -Path $excelFile 

# Function to create folders for IDs
function Create-FoldersForIDs {

    # Check if the ID column exists
    if ($data.ID) {
        # Create a folder for each ID
        foreach ($id in $data.ID) {
            $folderPath = Join-Path -Path $sheetName -ChildPath $id
            if (-not (Test-Path -Path $folderPath)) {
                New-Item -ItemType Directory -Path $folderPath | Out-Null
            }
        }
    } else {
        Write-Host "ID column does not exist in sheet $sheetName"
    }
}

# Create folders for each sheet in the list
foreach ($sheet in $sheets) {
    Create-FoldersForIDs -sheetName $sheet
}

