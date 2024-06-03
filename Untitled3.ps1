# Define the path to the Excel file
$excelFile = "C:\Project fot Beni\data.xlsx"

# Define the sheet names
$sheets = @("Sheet1", "Sheet2")  # Add sheet names as needed

# Function to create folders and save data
function Create-FoldersAndSaveData {

    # Import the data from the specified sheet
    $data = Import-Excel -Path $excelFile -NoHeader

    # Check if the ID column exists
    if ($data.PSObject.Properties.Name -contains 'ID') {
        # Create a folder for each unique ID
        foreach ($row in $data) {
            $id = [string]$row.ID
            $folderPath = "C:\Project fot Beni\Tests\test 5\folders\$id"  # Update with desired folder path

            # Create the main ID folder
            if (-not (Test-Path -Path $folderPath)) {
                New-Item -ItemType Directory -Path $folderPath | Out-Null
                Write-Host "Created folder: $folderPath"
            }

            # Create subfolders
            $subfolders = @("ק-0", "ק-1", "ק-2", "חוק")  # Update subfolder names
            foreach ($subfolder in $subfolders) {
                $subfolderPath = Join-Path -Path $folderPath -ChildPath $subfolder
                if (-not (Test-Path -Path $subfolderPath)) {
                    New-Item -ItemType Directory -Path $subfolderPath | Out-Null
                    Write-Host "Created subfolder: $subfolderPath"
                }
            }

            # Define mappings for URL and summary columns
            $mappings = @(
                @{ Subfolder = "ק-0"; UrlColumn = "C"; SummaryColumn = "B"; FileName = "K0" },
                @{ Subfolder = "ק-1"; UrlColumn = "E"; SummaryColumn = "D"; FileName = "K1" },
                @{ Subfolder = "ק-2"; UrlColumn = "G"; SummaryColumn = "F"; FileName = "K2" },
                @{ Subfolder = "חוק"; UrlColumn = "I"; SummaryColumn = "H"; FileName = "Law" }
            )

            # Process each mapping
            foreach ($mapping in $mappings) {
                if ($row.PSObject.Properties.Name -contains $mapping.UrlColumn -and $row.PSObject.Properties.Name -contains $mapping.SummaryColumn) {
                    $subfolderPath = Join-Path -Path $folderPath -ChildPath $mapping.Subfolder

                    # Download the file from URL
                    $url = $row."$($mapping.UrlColumn)"
                    if ($url -ne $null -and $url -ne "") {
                        $outputFilePath = Join-Path -Path $subfolderPath -ChildPath "$($mapping.FileName).pdf"
                        Download-File -url $url -outputPath $outputFilePath
                    }

                    # Save the summary content as a Word document
                    $summaryContent = $row."$($mapping.SummaryColumn)"
                    if ($summaryContent -ne $null -and $summaryContent -ne "") {
                        $wordFilePath = Join-Path -Path $subfolderPath -ChildPath "$($mapping.FileName)_Summary.docx"
                        Create-WordDocument -content $summaryContent -outputPath $wordFilePath
                    }
                }
            }
        }
    } else {
        Write-Host "ID column does not exist in sheet $sheetName"
    }
}
