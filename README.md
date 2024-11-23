# GoldSrc-Map-Check
Checks your maps and overview files for GoldSrc (Half-Life 1) engine games
GoldSrc Map Check 1.0

This is a quick and dirty script I made to check any GoldSrc game/mod folders to see if you are missing any .txt, .nav, .res, or overview file for any map (.bsp) for that game. I made this because my Counter-Strike folder was getting out of hand and trying to remember what maps I had overview and nav files for and which ones I didn't.

I also made a column for "Date modified" for the .bsp file so you can quickly check if you are worried if your version's date doesn't seem right (for example, your copy of de_rotterdam seems "too new" so you aren't sure you have the original file or if its a later modification, etc.

Due to all GoldSrc games using the same folder structures it should work with GoldSrc game/mod but I haven't tested many.

=========================

Installation and Running:
Step Zero: This only works on Windows (because PowerShell is a windows scripting language. Also, I only tested on Windows 11 but I would suspect it works on Windows 10 as well)

Step 0.5: You have to have a way to open Excel files. The ImportExcel function that the PowerShell script uses does not require Excel to be installed to run the script so if you don't have Excel you can still generate the file and then open the .xlsx with something like Google Sheets.

1) Unzip "map_check.bat" and "map_check.ps1" into the base folder of the game/mod you want to check your maps for.
	For example, Counter-Strike would be "cstrike" or even "cstrike_addon" or "cstrike_downloads" if you have custom content.
	NOT the maps or any other subfolders!
	
2) Run "map_check.bat" to help execute "map_check.ps1"
	The .bat file was made to help make running the PowerShell script easier. If you know how, you can run the .ps1 file directly, or even copy/paste the contents of the .ps1 file directly into a PowerShell terminal and run it, but the .bat should make it easier for anyone to run.
	
3) "map_check.bat" + "map_check.ps1" will create a new file in the directory you run in called "map_check.xlsx"
	"map_check.xlsx" is the file that will help you identify your maps and related files, open with Microsoft Excel, Google Sheets, etc.

4) Open "map_check.xlsx" to see your list of map files and if they have corresponding .txt, .nav, .res, and overview files and the date modified for the .bsp file. 
