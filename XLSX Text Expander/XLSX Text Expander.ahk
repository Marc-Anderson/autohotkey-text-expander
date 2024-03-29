#Requires AutoHotkey >=1.0 <1.9
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2

; name and path of file containing hotstrings
hotstringFilename := "hotstrings.xlsx"
hotstringFilepath := A_ScriptDir "\" hotstringFilename

; name of worksheet containing the hotstrings
hotstringWorksheetName := "Templates"

; name and path of file for keeping track of use for each hotstring
hotstringCounterFilename := "hotstring-counter.txt"
hotstringCounterFilepath := A_WorkingDir . "\" . hotstringCounterFilename

; name and path of file for splash image
splashImageFilename := "splashfile300x100.png"
splashImageFilepath := A_ScriptDir "\" splashImageFilename

; initialize the XL variable
XL :=

; show error if the hotstring file doesnt exist
if !FileExist(hotstringFilepath) {
    
    ; Alert the user that the file was not found
    MsgBox, The necessary workbook was not found. Please create a %hotstringFilename% file to continue.
    ExitApp
    Sleep, 2000
}

; create splash image for 
Gui, splash:New
Gui, +hwnd_SplashWindowId
Gui, splash:-Caption +alwaysontop
Gui, splash:Font, s14, Verdana
Gui, splash:Add, Text, x0 y50 w330 Center, AutoHotkey Text Expander
Gui, splash:Font, s8, Verdana
if FileExist(splashImageFilepath){
    Gui, Add, Picture, x15 y15 w300 h100, %splashImageFilepath%
}
Gui, splash:Add, Text, x275 y115 Center, Loading...
Gui, splash:Show, NoActivate w330 h130
WinSet, Transparent, 150, ahk_id %_SplashWindowId%

; try to load the hotstring file
try {
    ; Check if excel is active
    XL := ComObjectActive("Excel.Application")
} catch {
    ; If Excel is not active, create an instance
    XL := ComObjCreate("Excel.Application")
}
; MsgBox, % IsObject(XL) ; Is excel an object?
try {
    ; Make Excel invisible
    XL.Visible := 0
    ; Check if the workbook exists
    XL.Workbooks.Open(hotstringFilepath)
} catch {
    ; Make Excel invisible
    XL.Visible := 1
    ; Quit the application
    if(XL.Workbooks.Count = 0){
        XL.Application.Quit()
        XL := ""
    }

    ; remove the splash screen on error
    Gui, %_SplashWindowId%:Destroy

    ; Alert the user that there was an error opening the hotstring file
    MsgBox, Either the necessary workbook was not found or there was another error opening the hotstring file. Please create a %hotstringFilename% file and check your excel installation to continue.
    ExitApp
    Sleep, 2000
}

; object for keeping track of use for each hotstring
hotstringCounterObject := retrieveObjectFromFile(hotstringCounterFilepath)

; if the hotstring counter file exists keep track of hotstring use, otherwise do not save it
if FileExist(hotstringCounterFilepath) {

    ; save the hotstring counter before exiting
    OnExit(func("saveObjectRowsToTextFile").bind(hotstringCounterObject,hotstringCounterFilepath))
}

; select the sheet name containing the templates
hotstringWorksheet := XL.Worksheets(hotstringWorksheetName)

; activate the sheet with hotstrings
hotstringWorksheet.Activate

; Sort the data by column C so it can loop over all of the hotstrings without empty cells
hotstringWorksheet.UsedRange.Offset(1).Sort(XL.Columns(3), 1)

; Loop through all of the active cells
while(hotstringWorksheet.Range("C" . A_Index).Value != "") {

    ; Skip the first row as they are the header
    IfLess, A_Index, 2, Continue

    ; assign the value in column C of the current row to HotStringShortCut variable
    HotStringShortCut := hotstringWorksheet.Range("C" . A_Index).Value

    ; assign the value in column D of the current row to HotStringExtended variable
    HotStringExtended := hotstringWorksheet.Range("D" . A_Index).Value

    ; replaces any exclamation points and carriage returns with appropriate characters 
    HotStringExtended := StrReplace(StrReplace(StrReplace(HotStringExtended, "!","{!}"),"`r","{enter}"),A_Space A_Space,"{space 2}")

    ; assign both variables to a hotkey and execute via function
    hotstring(":*:" HotStringShortCut, func("executeHotstring").bind(HotStringExtended,hotstringCounterObject), 1)
    ; MsgBox, 4, , %HotStringShortCut% - %HotStringExtended%`n`nContinue?

}
; tell excel it's save so it wont harass you and close the document
XL.Application.Workbooks(hotstringFilename).saved := true
XL.Application.Workbooks(hotstringFilename).Close

if(XL.Workbooks.Count = 0){
    XL.Application.Quit()
}

; clear any unused variables
hotstringWorksheet := ""
HotStringExtended := ""
HotStringShortCut := ""
XL := ""

; add info about the application to task bar
AppInfoMenuVar := Func("AppInfoMenu")
Menu, Tray, Add, App Info, % AppInfoMenuVar

AppInfoMenu(){
    global
    Gui, AppInfo:New, +AlwaysOnTop
    Gui, AppInfo:+LabelAppInfoGui_On
    Gui, AppInfo:Font, s18, Verdana  
    Gui, AppInfo:Add, Text,, AutoHotkey Text Expander
    Gui, AppInfo:Font, s10, Verdana  
    Gui, AppInfo:Add, Text, w500 h200, This text expander allows you to automatically convert short phrases into long blocks of text. New shortcuts can be added in the included %hotstringFilename% file.`n`nExample, typing <ate will expand into "AutoHotkey Text Expander"`n`nBuilt In Hotstring:`n<now = DateTime(MM/dd/yyyy hh:mm:ss)`n`nIf you would like to keep track of how often your hotstrings are used, create a %hotstringCounterFilename% in the root folder where the application is stored and it will keep a running tally.
    Gui, AppInfo:Show,,AHK Text Expander Info
    ; MsgBox, You selected "%A_ThisMenuItem%" in menu "%A_ThisMenu%".
    return
}

; remove the splash screen on error
Gui, %_SplashWindowId%:Destroy

AppInfoGui_OnClose(GuiHwnd){
    Gui, %GuiHwnd%:Destroy
}

executeHotstring(HotStringExtendedText,hotstringCounterObject){

    ; check if extended text has <<input>> at the beginning, get input from user and replace all instances of <<template>>
    global
    if(SubStr(HotStringExtendedText, 1,9) == "<<input>>"){
        Sleep, 50
        createTextInputWindowAndWait()
        ; InputBox, UserInputValue, Text Replacement,,,180, 100
        HotStringExtendedText := StrReplace(HotStringExtendedText, "<<template>>", UserInputValue)
        HotStringExtendedText := SubStr(HotStringExtendedText, 10)
    }
    
    ; send hotstring ExtendedText
	sendinput % HotStringExtendedText

    ; update hotstring counter object
    thisHotkeyValue := StrReplace(A_ThisHotkey, ":*:")
    If (hotstringCounterObject.HasKey(thisHotkeyValue)){
        hotstringCounterObject[thisHotkeyValue]++
    } else {
        hotstringCounterObject[thisHotkeyValue] := 1
    }
    return
}

retrieveObjectFromFile(filename){
    tempObject := {}
    if FileExist(filename) {
        Loop {
            FileReadLine, line, %filename%, %A_Index%
            if ErrorLevel
                break
            lineData := StrSplit(line, ",")
            tempObject[lineData[1]] := lineData[2]
        }
    }
    return tempObject
}

saveObjectRowsToTextFile(targetObject, filepath){
    If (targetObject.Count() = 0){
        ; MsgBox, 0, DEBUG, No Data To Save, .5
        return
    }
    FileDelete, %filepath%
    combinedData := ""
    for keys, values in targetObject
    combinedData .= keys "," values "`n"
    FileAppend, %combinedData%, %filepath%
    ; MsgBox, 0, DEBUG, Data Saved To File`n`n %combinedDaata%, .5
}

createTextInputWindowAndWait(){
    global
    Gui, New, +HwndinputWindowID +AlwaysOnTop -Caption, Input Text
    Gui, Font, s14, Verdana
    Gui, Add, Edit, vUserInputValue x12 y10 w326
    Gui, Add, Button, Hidden default w0 h0, Submit
    Gui, Show, w350 h50
    WinWaitClose, Input Text
    ; return inputWindowID
    return
}

ButtonSubmit(){
    Gui, Submit
    Gui, Destroy
}

; built in hotstrings
:*:<now::
    FormatTime, CurrentDateTime,, MM/dd/yyyy hh:mm:ss 
    SendInput, %CurrentDateTime%
    return