# GoldSrc Map Check

## About

This is a basic script I made to check any GoldSrc game/mod folders to check the installed maps and catalog if they have .txt, .nav, .res, overview, and detail.txt files. I made this because my Counter-Strike folder was getting out of hand and trying to remember what maps I had overview and nav files for and which ones I didn't was becoming complicated.

I also made a column for "date modified" and "md5 checksum" for the .bsp file so you can try to determine if your version of a map file is the one you want it to be.

Due to all GoldSrc games using the same folder structures it should work with any GoldSrc game, but I've only tested a few.

If you have any questions, comments, or ideas for additional features, you can reach out to me on my [GameBanana Profile](https://gamebanana.com/members/2990169)

By using this software, you agree that I am NOT liable for any errors or data loss, or corruption that occur.

## Installation and Running

+ This only works on Windows (because PowerShell is a Windows scripting language. Also, I only tested on Windows 11 but I would suspect it works on Windows 10 as well)

+ You have to have a way to open Excel files. The ImportExcel function that the PowerShell script uses (which will also ask to intstall if not present and needs to be installed for this script to work) does not require Excel to be installed to run the script so if you don't have Excel you can still generate the file and then open the .xlsx file with something like Google Sheets.

1. Unzip the "GoldSrc Map Check" folder into your preferred directory.
	
2. Run "RunMapCheck.bat"
	- The .bat file was made to help make running the PowerShell script easier. If you know how, you can run the .ps1 file directly, or even copy/paste the contents of the .ps1 file directly into a PowerShell terminal and run it, but the .bat should make it easier for anyone to run.
	
3. The PowerShell script will create a new (or update an existing) Excel file in the "GoldSrc Map Check" folder named "Map Check.xlsx"
	- The terminal should tell you what it's doing each step of the way and prompt for certain actions.

4. Open the created Excel file to see your list of map files and if they have corresponding files and the date modified and md5 checksum for the .bsp file.
	- Open with Microsoft Excel, Google Sheets, etc.

![Google Sheets example](https://github.com/Lanecrest/GoldSrc-Map-Check/blob/main/screenshots/v1_2_preview.png)

## Updates

### 1.3 (11-26-24)

+ Added warning prompts, user prompts, and overall better and more information to the terminal
+ Updated folder structure and file names of this application
+ Added a column for detail.txt file presence
+ Version info, disclaimers, updated language, etc

### 1.2 (11-24-24)

+ User friendly error handling with more information displayed in the terminal as the script runs.
+ Columns will now simply say "yes" if a .nav, .txt, or .res is present, and if it has an overview will just state if it is a .bmp or a .tga
+ Added a checksum column to better help identify if your version of a map is the version you are tying to make sure you have.

### 1.1 (11-23-24)

+ Updated so that the script no longer need to run in your Half-Life game/mod folders, and can be ran from anywhere.
+ It will try to autodetect your Steam install of Half-Life to make navigating to the directory quicker.
+ The file will create a worksheet based on the GoldSrc game/mod parent directory you select in an excel workbook.
+ The file will generate in the directory you run the script in, and you can run it multiple times, meaning you can have multiple worksheets in the same workbook for each of your GoldSrc games.
+ The map names will write in alphabetical order by default, with filters.
+ Better error handling and better comments in the code.

### 1.0 (11-23-24)

+ Initial release, uses powershell to create an excel (xlsx) file to check map files in a GoldSrc game/mod folder. Additional features may or may not be implemented in the future. See screenshots for example on what the generated file looks like.

### Credits

Lanecrest Tech © 2024
Doctor Worm

This program is free software released under the MIT License.