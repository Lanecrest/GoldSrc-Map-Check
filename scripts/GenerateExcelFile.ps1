# Check if the Excel file already exists
$fileName = "Map Check.xlsx"
$excelPath = ".\$fileName"
if (Test-Path $excelPath) {
    Write-Host "`nNotice: The file '$fileName' already exists and will be overwritten." -ForegroundColor Yellow
}

# Install the ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
	Write-Host "Notice: ImportExcel module is required in order for script to work. Looking for package..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
	Write-Host "`nPackage installed. Continuing..."
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

# Function to search for a target folder structure
function Find-Folder {
	$targetFolder = "SteamLibrary\steamapps\common\Half-Life"
    $drives = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root
    foreach ($drive in $drives) {
        $possiblePath = Join-Path -Path $drive -ChildPath $targetFolder
        if (Test-Path $possiblePath) {
            return $possiblePath
        }
    }
    return $null
}

# Create a folder browser dialog object that attempts to start at the target folder
Add-Type -AssemblyName System.Windows.Forms
Write-Host "`nLooking for Half-Life Steam installation folder..."
$defaultPath = Find-Folder
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select a GoldSrc game folder containing a 'maps' subfolder"

# Set the default path if the target folder was found
if ($defaultPath) {
    $folderBrowser.SelectedPath = $defaultPath
    Write-Host "`nFolder found: $defaultPath"
}
else {
    Write-Host "`nFolder not found. Navigate to your Half-Life installation folder."
}

# Display folder browser dialog for user folder selection
if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $rootFolder = $folderBrowser.SelectedPath
	$folderName = Split-Path -Path $rootFolder -Leaf
    Write-Host "`nSelected game folder: $folderName"
}
else {
    Write-Host "`nNo folder selected. Exiting."
    exit
}

# Check if a worksheet for the selected game already exists in the file
try {
    $worksheet = Import-Excel -Path $excelPath -WorksheetName $folderName
    Write-Host "`nNotice: A worksheet for '$folderName' already exists and will be overwritten." -ForegroundColor Yellow
	# Prompt user to continue if the worksheet is found
    $userInput = Read-Host "`nDo you want to continue? (Y/N)"
	if ($userInput -in @("Y", "y")) {
    }
	else {
        Write-Host "`nExiting."
        exit
	}
}
catch {
	Write-Host "`nData will be written to a new worksheet named '$folderName' in '$fileName'"
}

# Define the paths to the "maps" and "overviews" subfolders
$mapsFolder = Join-Path -Path $rootFolder -ChildPath "maps"
$overviewsFolder = Join-Path -Path $rootFolder -ChildPath "overviews"

# Check if the "maps" folder exists
if (-not (Test-Path -Path $mapsFolder)) {
    Write-Host "`nError: The 'maps' folder does not exist in the current directory. Exiting." -ForegroundColor Red
	exit
}

# Check if the "overviews" folder exists
if (-not (Test-Path -Path $overviewsFolder)) {
    Write-Host "`nNotice: The 'overviews' folder does not exist in the current directory. Continuing..." -ForegroundColor Yellow
}
# Suppress errors if 'overviews' folder doesn't exist or is empty
else {
    $overviewFiles = Get-ChildItem -Path $overviewsFolder -File -ErrorAction SilentlyContinue
}

# Get the files in the "maps" folder
$mapFiles = Get-ChildItem -Path $mapsFolder -File

# Create a hashtable for storing data
$fileData = @{}

# Populate the hashtable with .bsp files with their base name as the key
Write-Host "`nAnalyzing folders...`n"
foreach ($file in $mapFiles) {
    if ($file.Extension -eq ".bsp") {
        $baseName = $file.BaseName
        $fileData[$baseName] = @{
            "bsp" = $file.BaseName
            "txt" = ""
            "nav" = ""
            "res" = ""
            "overview" = ""
			"detail" = ""
            "dateModified" = $file.LastWriteTime
			"md5" = Get-FileChecksum -filePath $file.FullName
		}
	Write-Host "Processing file: $file" -ForegroundColor Cyan
    }
}

# Check the "maps" folder for .txt, .nav, .res, and detail files that match the found .bsp files
foreach ($file in $mapFiles) {
    $baseName = $file.BaseName
    if ($file.Extension -in @(".txt", ".nav", ".res")) {
        if ($fileData.ContainsKey($baseName)) {
			Write-Host "Processing file: $file" -ForegroundColor Cyan
            $fileData[$baseName][$file.Extension.TrimStart(".")] = "yes"
        }
    }
	if ($file.Extension -eq ".txt" -and $file.BaseName.EndsWith("_detail")) {
		$detailName = $file.BaseName -replace "_detail$"
		if ($fileData.ContainsKey($detailName)) {
			Write-Host "Processing file: $file" -ForegroundColor Cyan
			$fileData[$detailName]["detail"] = "yes"
		}
	}
}

# Check the "overviews" folder for .bmp and .tga files that match the found .bsp files
foreach ($file in $overviewFiles) {
    $baseName = $file.BaseName
    if ($file.Extension -in @(".bmp", ".tga")) {
        if ($fileData.ContainsKey($baseName)) {
			Write-Host "Processing file: $file" -ForegroundColor Cyan
            $fileData[$baseName]["overview"] = $file.Extension.TrimStart(".")
        }
    }
}

# Convert the hashtable to a list of custom objects for Excel export
Write-Host "`nWriting excel file..."
$table = $fileData.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        "Map Name" = $_.Value["bsp"]
        ".txt" = $_.Value["txt"]
        ".nav" = $_.Value["nav"]
        ".res" = $_.Value["res"]
        "Overview" = $_.Value["overview"]
		"Detail" = $_.Value["detail"]
        "Date Modified" = $_.Value["dateModified"]
		"MD5 Checksum" = $_.Value["md5"]
    }
} | Sort-Object -Property "Map Name"

# Export the data to an Excel file
$retry = $true
while ($retry) {
	try {
		$table | Export-Excel -Path .\$fileName -WorksheetName $folderName -AutoSize -AutoFilter
		Write-Host "`nData successfully written to worksheet '$folderName' in file '$fileName'" -ForegroundColor Green
		$retry = $false
		# Prompt user to open the file
		$userInput = Read-Host "`nDo you want to open the file? (Y/N)"
		if ($userInput -in @("Y", "y")) {
			Start-Process -FilePath .\$fileName
		}
	}
	catch {
		# Error handling for the file likely being open
		if ($_ -match "Could not open Excel Package") {
			Write-Host "`nError: The file '$fileName' could not be written to. If it is open, please close the file and try again." -ForegroundColor Red
		}
		# Error handling for other errors
		else {
			Write-Host "`nAn error occurred while creating the Excel file: $_" -ForegroundColor Red
		}
		# Prompt user for retry
        $userInput = Read-Host "`nDo you want to try again? (Y/N)"
        if ($userInput -notin @("Y", "y")) {
            Write-Host "`nExiting."
            $retry = $false
        }
	}
}