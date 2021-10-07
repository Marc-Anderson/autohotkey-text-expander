# AutoHotkey-Text-Expander
A lightweight spreadsheet(CSV,XLSX) based text expander built in AutoHotkey.

## What it does
This text expander allows you to automatically convert short phrases into long blocks of text. For example, typing <ate will expand into "AutoHotkey Text Expander" or typing <now will expand into the date and time formatted like this. MM/dd/yyyy hh:mm:ss. New shortcuts can be added in the hotstrings.xlsx(or csv) file.

## How to setup
There is no installation process. Download the exe and hostrings.xlsx(or csv) and put them in the same folder, launch the app and it will read the hostrings file for all of your shortcuts. It doesn't continuously read the hotstrings file so if you add new ones you need to close the app and relaunch. 

## FAQ
* There is nothing special about the hostrings file. Download the supplied one or make your own so long as the sheet, document names and columns are the same as the original.
* If you already have ahk installed you can use the ahk file instead of the exe. 
* I used the xlsx version every day for work.
* I use the < at the beginning of all of my hotstrings but thats not necessary. You cant change the <now hotstring but any new ones you add can use any prefix or none at all. 
