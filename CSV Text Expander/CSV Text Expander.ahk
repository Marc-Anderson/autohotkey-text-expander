#Requires AutoHotkey >=1.0 <1.9
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2

; name and path of file containing hotstrings
hotstringFilename := "hotstrings.csv"
hotstringFilepath := A_ScriptDir "\" hotstringFilename

; name and path of file for keeping track of use for each hotstring
hotstringCounterFilename := "hotstring-counter.txt"
hotstringCounterFilepath := A_WorkingDir . "\" . hotstringCounterFilename

; show error if the hotstring file doesnt exist
if !FileExist(hotstringFilepath) {
    
    ; Alert the user that the file was not found
    MsgBox, The necessary hotstring file was not found. Please create a %hotstringFilename% file to continue.
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

FileRead, CSV, %hotstringFilepath%
Loop, Parse, CSV, `r, `n 
{
    ; save the line number to a variable
    LineNumber := A_Index -1

    ; skip the first row of content
    IfEqual, A_Index, 1, Continue

    Loop, Parse, A_LoopField, CSV
    {
    ; MsgBox, 4, , %A_Index% is:`n%A_LoopField%`nContinue?
        ; if the cell number is odd, assign the current cell to the HotStringShortCut variable
        if ( Mod(A_Index, 2) != 0) {
            HotStringShortCut := A_LoopField ;else
            ; MsgBox, 4, , %LineNumber%-%A_Index% is:`n%A_LoopField%`n%HotStringShortCut%`nContinue?

        ; if the cell number is even, assign the current cell to the HotStringExtended variable
        } else {
            HotStringExtended := A_LoopField
            ; MsgBox, 4, , %LineNumber%-%A_Index% is:`n%A_LoopField%`n%HotStringExtended%`nContinue?
        }

        ; if the cell number is even, assign both variables to a hotkey
        if ( Mod(A_Index, 2) = 0) {
            ; MsgBox, 4, , %HotStringShortCut% - %HotStringExtended%`n`nContinue?
            HotStringExtended := StrReplace(StrReplace(StrReplace(HotStringExtended, "!","{!}"),"`r","{enter}"),A_Space A_Space,"{space 2}")
            ; assign both variables to a hotkey and execute via function
            hotstring(":*:" HotStringShortCut, func("executeHotstring").bind(HotStringExtended,hotstringCounterObject), 1)
        }

        ; %LineNumber% is the current row
        ; %A_Index% is the current column
        ; %A_LoopField% is the current field
        
        ; MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
        ; IfMsgBox, No
        ;     return
    }
}

executeHotstring(HotStringExtendedText,hotstringCounterObject){

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

; add info about the application to task bar
AppInfoMenuVar := Func("AppInfoMenu")
Menu, Tray, Add, App Info, % AppInfoMenuVar

AppInfoMenu(){
    global
    Gui, ateInfo:New, +AlwaysOnTop
    Gui, ateInfo:Font, s18, Verdana  
    Gui, ateInfo:Add, Text,, AutoHotkey Text Expander
    Gui, ateInfo:Font, s10, Verdana  
    Gui, ateInfo:Add, Text, w500 h200, This text expander allows you to automatically convert short phrases into long blocks of text. New shortcuts can be added in the included %hotstringFilename% file.`n`nExample, typing <ate will expand into "AutoHotkey Text Expander"`n`nBuilt In Hotstring:`n<now = DateTime(MM/dd/yyyy hh:mm:ss)`n`nIf you would like to keep track of how often your hotstrings are used, create a %hotstringCounterFilename% in the root folder where the application is stored and it will keep a running tally.
    Gui, ateInfo:Show,,AHK Text Expander Info
    ; MsgBox, You selected "%A_ThisMenuItem%" in menu "%A_ThisMenu%".
    return
}

; built in hotstrings
:*:<now::
    FormatTime, CurrentDateTime,, MM/dd/yyyy hh:mm:ss 
    SendInput, %CurrentDateTime%
    return