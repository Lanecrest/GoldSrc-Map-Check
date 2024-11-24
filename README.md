# GoldSrc Map Check

## About

This is a quick and dirty script I made to check any GoldSrc game/mod folders to see if you are missing any .txt, .nav, .res, or overview file for any map (.bsp) for that game. I made this because my Counter-Strike folder was getting out of hand and trying to remember what maps I had overview and nav files for and which ones I didn't was becoming complicated.

I also made a column for "date modified" and "md5 checksum" for the .bsp file so you can try to determine if your version of a map file is the one you need it to be.

Due to all GoldSrc games using the same folder structures it should work with any GoldSrc game, but I haven't tested many.

If you have any questions, comments, or ideas for additional features, you can reach out to me on my [GameBanana Profile](https://gamebanana.com/members/2990169)

## Installation and Running

+ This only works on Windows (because PowerShell is a windows scripting language. Also, I only tested on Windows 11 but I would suspect it works on Windows 10 as well)

+ You have to have a way to open Excel files. The ImportExcel function that the PowerShell script uses (which will also ask to intstall if not present and needs to be installed for this script to work) does not require Excel to be installed to run the script so if you don't have Excel you can still generate the file and then open the .xlsx with something like Google Sheets.

1. Unzip "map_check.bat" and "map_check.ps1" into a preferred directory on your computer.
	
2. Run "map_check.bat" to help execute "map_check.ps1"
	- The .bat file was made to help make running the PowerShell script easier. If you know how, you can run the .ps1 file directly, or even copy/paste the contents of the .ps1 file directly into a PowerShell terminal and run it, but the .bat should make it easier for anyone to run.
	
3. "map_check.ps1" will create a new (or update an existing) excel file in the directory you run it in named "map_check.xlsx"
	- The terminal should tell you what it's doing each step of the way.

4. Open "map_check.xlsx" to see your list of map files and if they have corresponding .txt, .nav, .res, and overview files and the date modified and md5 checksum for the .bsp file.
	- This generated excel file is the file that will help you identify your maps and related files. Open with Microsoft Excel, Google Sheets, etc.

![Google Sheets example](https://github.com/Lanecrest/GoldSrc-Map-Check/blob/main/screenshots/v1_2_preview.png)

## Updates

### 1.2 (11-24-24)

+ User friendly error handling with more information displayed in the terminal as the script runs.
+ Colums will now simply say "yes" if a .nav, .txt, or .res is present, and if it has an overview will just state if it is a .bmp or a .tga
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