# Install the ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
	Write-Host "Notice: ImportExcel module is required in order for script to work." -ForegroundColor Red
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

# Function to calculate the MD5 checksum of a file
function Get-FileChecksum {
    param (
        [string]$filePath
    )
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $stream = [System.IO.File]::OpenRead($filePath)
    $checksumBytes = $md5.ComputeHash($stream)
    $stream.Close()
    return [BitConverter]::ToString($checksumBytes) -replace '-'
}

# Search for the default target folder structure and set the default path for the FolderBrowserDialog
$defaultPath = Find-Folder

# Create and configure the folder browser dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select a game folder containing 'maps' and 'overviews'."

# Set the default path if the target folder was found
if ($defaultPath) {
    $folderBrowser.SelectedPath = $defaultPath
    Write-Host "Default path found: $defaultPath"
}
else {
    Write-Host "Target folder structure not found. Navigate to your Half-Life install folder."
}
Write-Host "Select the game folder containing the desired 'maps' and 'overviews' subfolders."

# Show the browser dialog and starting path
if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $rootFolder = $folderBrowser.SelectedPath
    Write-Host "Selected folder: $rootFolder"
}
else {
    Write-Host "No folder selected. Exiting."
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
    Write-Host "Warning: The 'overviews' folder does not exist in the current directory. Continuing." -ForegroundColor Red
}
else {
    # Suppress errors if 'overviews' folder doesn't exist or is empty
    $overviewFiles = Get-ChildItem -Path $overviewsFolder -File -ErrorAction SilentlyContinue
}

# Get the files in the "maps" folder
$mapFiles = Get-ChildItem -Path $mapsFolder -File

# Create a hashtable to store the data with the base file name as the key
Write-Host "Analyzing folders..."
$fileData = @{}

# Populate the hashtable with .bsp files as the main key
foreach ($file in $mapFiles) {
    if ($file.Extension -eq ".bsp") {
        $baseName = $file.BaseName
        $fileData[$baseName] = @{
            "bsp" = $file.BaseName
            "txt" = ""
            "nav" = ""
            "res" = ""
            "overview" = ""
            "dateModified" = $file.LastWriteTime
			"md5" = Get-FileChecksum -filePath $file.FullName
		}
    }
}

# Match other files (.txt, .nav, .res) to the .bsp files based on the base name
foreach ($file in $mapFiles) {
    $baseName = $file.BaseName
    if ($file.Extension -in @(".txt", ".nav", ".res")) {
        if ($fileData.ContainsKey($baseName)) {
            $fileData[$baseName][$file.Extension.TrimStart(".")] = "yes"
        }
    }
}

# Check the "overviews" folder for .bmp and .tga files that match the .bsp files based on the base name
foreach ($file in $overviewFiles) {
    $baseName = $file.BaseName
    if ($file.Extension -in @(".bmp", ".tga")) {
        if ($fileData.ContainsKey($baseName)) {
            $fileData[$baseName]["overview"] = $file.Extension
        }
    }
}

# Convert the hashtable to a list of custom objects for Excel export
Write-Host "Writing excel file..."
$table = $fileData.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        "Map" = $_.Value["bsp"]
        ".txt" = $_.Value["txt"]
        ".nav" = $_.Value["nav"]
        ".res" = $_.Value["res"]
        "Overview" = $_.Value["overview"]
        "Date Modified" = $_.Value["dateModified"]
		"MD5 Checksum" = $_.Value["md5"]
    }
} | Sort-Object -Property "Map"

# Export the data to an Excel file
$workbookName = "map_check.xlsx"
$worksheetName = Split-Path -Path $rootFolder -Leaf
try {
    $table | Export-Excel -Path .\$workbookName -WorksheetName $worksheetName -AutoSize -AutoFilter
    Write-Host "Excel file written successfully: $workbookName"
}
catch {
    # If the file is open or any other issue occurs, handle it gracefully
    if ($_ -match "Could not open Excel Package") {
        Write-Host "Error: The file '$workbookName' could not be written to. If it is open, please close the file and try again." -ForegroundColor Red
    }
	else {
        # Generic error handling for other errors
        Write-Host "An error occurred while creating the Excel file: $_" -ForegroundColor Red
    }
}