# AutoHotkey-Text-Expander
A lightweight spreadsheet(CSV,XLSX) based text expander built in AutoHotkey.

## What it does
This text expander allows you to automatically convert short phrases into long blocks of text. For example, typing <ate will expand into "AutoHotkey Text Expander" or typing <now will expand into the date and time formatted like this. MM/dd/yyyy hh:mm:ss. New shortcuts can be added in the hotstrings.xlsx(or csv) file.

## How to setup
There is no installation process. Download the exe and hotstrings.xlsx(or csv) and put them in the same folder, launch the app and it will read the hotstrings file for all of your shortcuts. It doesn't continuously read the hotstrings file so if you add new ones you need to close the app and relaunch. When the app is launched there is an icon in the taskbar near the clock where you can exit or pause the app.

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
| ID | Name         | HotString | Extended Text                                                                                                                         |
| -- | ------------ | --------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| 0  | My Name      | <ate      | AutoHotkey Text Expander                                                                                                              |
| 1  | Greeting     |           | Hello World!                                                                                                                          |
| 2  | Grocery Note | <list     | This is the list of items I need from the store.<br><br>    \* Apples<br>    \* Oranges<br>    \* Paper Towels<br><br>That’s it, the end of the list. |
| 3  | Ramble       | <blah     | blah blah blah |
  
  
   
  
## CSV hotstring file format
  
Filename: hotstrings.csv
  
### Table
| HotString | Extended Text |
| --------- | ------------- |
| <ate      | AutoHotkey Text Expander |
| <sorry    | I am sorry for the inconvenience. |
| <hi       | Hello World |
| <list     | This is the list of items I need from the store.<br><br> \* Apples<br> \* Oranges<br> \* Paper Towels<br><br> That’s it, the end of the list. |
| test      | This is a hotstring without a prefix. |
