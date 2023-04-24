# AutoHotkey-Text-Expander
A lightweight spreadsheet(CSV,XLSX) based text expander built in AutoHotkey v1.x.

## What it does
This text expander allows you to automatically convert short phrases into long blocks of text. For example, typing <ate will expand into "AutoHotkey Text Expander" or typing <now will expand into the date and time formatted like this. MM/dd/yyyy hh:mm:ss. New shortcuts can be added in the hotstrings.xlsx(or csv) file.

## How to setup
There is no installation process. Download the exe and hotstrings.xlsx(or csv) and put them in the same folder, launch the app and it will read the hotstrings file for all of your shortcuts. It doesn't continuously read the hotstrings file so if you add new ones you need to close the app and relaunch. When the app is launched there is an icon in the taskbar near the clock where you can exit or pause the app.

## Options

### Hotstring Counter Feature
By adding a `hotstring-counter.txt` file to the root directory, you can enable the Hotstring Counter feature. This feature will keep track of the number of times each hotstring is used.

1. Create a plain text file named `hotstring-counter.txt` in the root directory.
2. Each time a hotstring is used, the counter will be incremented for that specific hotstring.
3. The results will be saved to the `hotstring-counter.txt` file when the app is closed.
4. The `hotstring-counter.txt` file will store the hotstring usage data in csv format. 
5. Renaming the file from `.txt` to `.csv` will allow you to open it in your favorite spreadsheet editor. 

### Splash Screen Feature(xlsx only)
By adding a `splashfile300x100.png` image to the root folder, you can enable the Splash Screen feature. This feature will display a custom image while the script is loading hotstrings from the Excel file.
1. Create a 300x100 pixels image (PNG format) with your desired design.
2. Save the image as `splashfile300x100.png` in the root folder.
3. The `splashfile300x100.png` image will be displayed as a splash screen while the script is loading hotstrings from the Excel file.
4. Once the hotstrings are loaded, the splash screen will disappear, and the app will be ready for use.

### Input Replacement Feature
By adding `<<input>>` to the beginning of any extended text and `<<template>>` elsewhere in the extended text, the user will be prompted for input when the hotstring is executed. Upon pressing enter all instances of `<<template>>` will be replaced with the users input.

## FAQ
* There is nothing special about the hotstrings file. Download the supplied one or make your own so long as the sheet, document names and columns are the same as the original.
* If you already have ahk installed you can use the ahk file instead of the exe. 
* I use the xlsx version every day for work.
* I use the < at the beginning of all of my hotstrings but thats not necessary. You cant change the <now hotstring but any new ones you add can use any prefix or none at all.
* This is just a personal project because I wanted one and my boss would never pay for something like this but if you have suggestions I'd be open to adding/changing stuff. 

## Excel hotstring file format

Filename: hotstrings.xlsx  
Sheetname: Templates  

### Table
| ID | Name | HotString | Extended Text |
| -- | ---- | --------- | ------------- |
| 0 | Introduction | <ate | AutoHotkey Text Expander |
| 1 | Apology | <sorry | I am sorry for the inconvenience. |
| 2 | Greeting | <hi | Hello World |
| 3 | Grocery Note | <list | This is the list of items I need from the store.<br><br>    \* Apples<br>    \* Oranges<br>    \* Paper Towels<br><br>That’s it, the end of the list. |
| 4 | Test NoPrefix | test | This is a hotstring without a prefix. |
| 5 | Template Replace Input | <input | `<<input>>`This message is for `<<template>>`. We are trying to reach `<<template>>` about their cars extended warranty. |


## CSV hotstring file format

Filename: hotstrings.csv

### Table
| HotString | Extended Text |
| --------- | ------------- |
| <ate | AutoHotkey Text Expander |
| <sorry | I am sorry for the inconvenience. |
| <hi | Hello World |
| <list | This is the list of items I need from the store.<br><br>    \* Apples<br>    \* Oranges<br>    \* Paper Towels<br><br>That’s it, the end of the list. |
| test | This is a hotstring without a prefix. |
| <input | `<<input>>`This message is for `<<template>>`. We are trying to reach `<<template>>` about their cars extended warranty. |
