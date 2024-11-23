# Install the ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Get the path to the "maps" and "overviews" folders (assuming both are in the current working directory)
$mapsFolder = ".\maps"
$overviewsFolder = ".\overviews"

# Check if the "maps" folder exists
if (-not (Test-Path -Path $mapsFolder)) {
    Write-Host "The 'maps' folder does not exist in the current directory."
    exit
}

# Check if the "overviews" folder exists
if (-not (Test-Path -Path $overviewsFolder)) {
    Write-Host "The 'overviews' folder does not exist in the current directory."
    exit
}

# Get the files in the "maps" folder
$mapFiles = Get-ChildItem -Path $mapsFolder -File

# Create a hashtable to store the data, with the base file name as the key
$fileData = @{}

# Populate the hashtable with .bsp files as the main key
foreach ($file in $mapFiles) {
    if ($file.Extension -eq ".bsp") {
        $baseName = $file.BaseName
        $fileData[$baseName] = @{
            "bsp" = $file.Name
            "txt" = ""
            "nav" = ""
            "res" = ""
            "overview" = ""  # Add overview column
            "dateModified" = $file.LastWriteTime  # Add Date Modified column
        }
    }
}

# Now, match the other files (.txt, .nav, .res) to the .bsp files based on the base name
foreach ($file in $mapFiles) {
    $baseName = $file.BaseName
    if ($file.Extension -in @(".txt", ".nav", ".res")) {
        if ($fileData.ContainsKey($baseName)) {
            $fileData[$baseName][$file.Extension.TrimStart(".")] = $file.Name
        }
    }
}

# Now, check the "overviews" folder for .bmp and .tga files that match .bsp names
$overviewFiles = Get-ChildItem -Path $overviewsFolder -File

foreach ($file in $overviewFiles) {
    $baseName = $file.BaseName
    if ($file.Extension -in @(".bmp", ".tga")) {
        if ($fileData.ContainsKey($baseName)) {
            $fileData[$baseName]["overview"] = $file.Name
        }
    }
}

# Convert the hashtable to a list of custom objects for Excel export
$table = $fileData.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        ".bsp" = $_.Value["bsp"]
        ".txt" = $_.Value["txt"]
        ".nav" = $_.Value["nav"]
        ".res" = $_.Value["res"]
        "overview" = $_.Value["overview"]
        ".bsp modified" = $_.Value["dateModified"]  # Add Date Modified to the output for .bsp file
    }
}

# Export the data to an Excel file
$table | Export-Excel -Path .\map_check.xlsx -WorksheetName "Files" -AutoSize
Write-Host "Excel file created successfully: map_check.xlsx"
