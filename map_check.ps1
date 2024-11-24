# Install the ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Add support for Windows Forms to use the FolderBrowserDialog
Add-Type -AssemblyName System.Windows.Forms

# Define the default target folder structure
$targetFolderStructure = "SteamLibrary\steamapps\common\Half-Life"

# Function to search for the target folder structure
function Find-Folder {
    $drives = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root
    foreach ($drive in $drives) {
        $possiblePath = Join-Path -Path $drive -ChildPath $targetFolderStructure
        if (Test-Path $possiblePath) {
            return $possiblePath
        }
    }
    return $null
}

# Search for the default target folder structure and set the default path for the FolderBrowserDialog
$defaultPath = Find-Folder

# Create and configure the folder browser dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the parent folder containing 'maps' and 'overviews'."

# Set the default path if the target folder was found
if ($defaultPath) {
    $folderBrowser.SelectedPath = $defaultPath
    Write-Host "Default path found: $defaultPath"
}
else {
    Write-Host "Target folder structure not found. Navigate to your Half-Life install folder."
}
Write-Host "Select the parent folder containing the desired 'maps' and 'overviews' subfolders."

# Show the browser dialog and starting path
if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $rootFolder = $folderBrowser.SelectedPath
    Write-Host "Selected folder: $rootFolder"
}
else {
    Write-Host "No folder selected. Exiting." -ForegroundColor Red
    exit
}

# Define the paths to the "maps" and "overviews" subfolders
$mapsFolder = Join-Path -Path $rootFolder -ChildPath "maps"
$overviewsFolder = Join-Path -Path $rootFolder -ChildPath "overviews"

# Check if the "maps" folder exists
if (-not (Test-Path -Path $mapsFolder)) {
    Write-Host "Error: The 'maps' folder does not exist in the current directory. Exiting." -ForegroundColor Red
	exit
}

# Check if the "overviews" folder exists
if (-not (Test-Path -Path $overviewsFolder)) {
    Write-Host "Error: The 'overviews' folder does not exist in the current directory. Continuing." -ForegroundColor Red
}
else {
    # Suppress errors if 'overviews' folder doesn't exist or is empty
    $overviewFiles = Get-ChildItem -Path $overviewsFolder -File -ErrorAction SilentlyContinue
}

# Get the files in the "maps" folder
$mapFiles = Get-ChildItem -Path $mapsFolder -File

# Create a hashtable to store the data with the base file name as the key
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
            "overview" = ""
            "dateModified" = $file.LastWriteTime
        }
    }
}

# Match other files (.txt, .nav, .res) to the .bsp files based on the base name
foreach ($file in $mapFiles) {
    $baseName = $file.BaseName
    if ($file.Extension -in @(".txt", ".nav", ".res")) {
        if ($fileData.ContainsKey($baseName)) {
            $fileData[$baseName][$file.Extension.TrimStart(".")] = $file.Name
        }
    }
}

# Check the "overviews" folder for .bmp and .tga files that match the .bsp files based on the base name
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
        ".bsp modified" = $_.Value["dateModified"]
    }
} | Sort-Object -Property ".bsp"  # Sort the data by the .bsp column

# Export the data to an Excel file
$worksheetName = Split-Path -Path $rootFolder -Leaf
$table | Export-Excel -Path .\map_check.xlsx -WorksheetName $worksheetName -AutoSize -AutoFilter
Write-Host "Excel file created successfully: map_check.xlsx"
