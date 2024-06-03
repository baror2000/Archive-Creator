# Install required module: ImportExcel
Install-Module ImportExcel

# Import the module
Import-Module ImportExcel

# Read Excel data
$excelData = Import-Excel -Path "C:\Project fot Beni\Data.xlsx"

# Loop through each row in the Excel file
foreach ($row in $excelData) {
    # Create a folder for each ID
    $folderPath = "C:\Project fot Beni\test6\Folders\$($row.ID)"
    New-Item -ItemType Directory -Path $folderPath -Force | Out-Null

    # Create subfolders "ק-0 ק-1 ק-2 חוק" within the folder
    foreach ($subFolder in ("ק-0", "ק-1", "ק-2", "חוק")) {
        $subFolderPath = Join-Path -Path $folderPath -ChildPath $subFolder
        New-Item -ItemType Directory -Path $subFolderPath -Force | Out-Null
    }

    # Download files from URLs and create PDF files
    Invoke-WebRequest -Uri $row.C -OutFile "$folderPath\ק-0\$($row.ID)_ק-0.docx"
    Invoke-WebRequest -Uri $row.E -OutFile "$folderPath\ק-1\$($row.ID)_ק-1.pdf"
    Invoke-WebRequest -Uri $row.F -OutFile "$folderPath\ק-2\$($row.ID)_ק-2.pdf"
    Invoke-WebRequest -Uri $row.I -OutFile "$folderPath\חוק\$($row.ID)_חוק.docx"
}
