#Requires AutoHotkey >=2.0
; #NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode("Input")  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir(A_ScriptDir)  ; Ensures a consistent starting directory.
SetTitleMatchMode(2)

; name and path of file containing hotstrings
hotstringFilename := "hotstrings.csv"
hotstringFilepath := A_ScriptDir "\" hotstringFilename

; name and path of file for keeping track of use for each hotstring
hotstringCounterFilename := "hotstring-counter.txt"
hotstringCounterFilepath := A_WorkingDir "\" hotstringCounterFilename

; show error if the hotstring file doesnt exist
if !FileExist(hotstringFilepath) {
    
    ; Alert the user that the file was not found
    MsgBox("The necessary hotstring file was not found. Please create a " hotstringFilename " file to continue.")
    ExitApp()
    Sleep(2000)
}

; object for keeping track of use for each hotstring
hotstringCounterObject := retrieveObjectFromFile(hotstringCounterFilepath)

; if the hotstring counter file exists keep track of hotstring use, otherwise do not save it
if FileExist(hotstringCounterFilepath) {

    ; save the hotstring counter before exiting
    OnExit(saveObjectRowsToTextFile.Bind(hotstringCounterObject, hotstringCounterFilepath))
}

; MsgBox("the filepath is " hotstringFilepath)

CSVContents := FileRead(hotstringFilepath)

Loop parse, CSVContents, "`r", "`n" 
{
    ; save the line number to a variable
    LineNumber := A_Index -1

    ; Skip the first row as they are the header
    if (A_Index < 2) {
        continue
    }

    ; MsgBox("reading file:" hotstringFilepath "`n`nrow: " )

    Loop parse, A_LoopField, "CSV"
    {
        ; MsgBox(A_Index " is:`n" A_LoopField "`nContinue?")
        ; if the cell number is odd, assign the current cell to the HotStringShortCut variable
        if ( Mod(A_Index, 2) != 0) {
            HotStringShortCut := A_LoopField ;else
            ; MsgBox(4, , LineNumber "-" A_Index " is:`n" A_LoopField "`n" HotStringShortCut "`nContinue?")

        ; if the cell number is even, assign the current cell to the HotStringExtended variable
        } else {
            HotStringExtended := A_LoopField
            ; MsgBox(4, , LineNumber "-" A_Index " is:`n" A_LoopField "`n" HotStringExtended "`nContinue?")
        }

        ; if the cell number is even, assign both variables to a hotkey
        if ( Mod(A_Index, 2) = 0) {
            ; MsgBox, 4, , %HotStringShortCut% - %HotStringExtended%`n`nContinue?
            HotStringExtended := StrReplace(StrReplace(StrReplace(HotStringExtended, "!","{!}"),"`r","{enter}"),A_Space A_Space,"{space 2}")

            ; assign both variables to a hotkey and execute via function
            Hotstring(":*:" HotStringShortCut, executeHotstring.Bind(HotStringExtended, hotstringCounterObject), "On")
            ; MsgBox("HotStringShortCut: " HotStringShortCut "`nHotStringExtended: " HotStringExtended "`n`nContinue?")

        }

        ; %LineNumber% is the current row
        ; %A_Index% is the current column
        ; %A_LoopField% is the current field

        ; MsgBox(4, , "Field " LineNumber "-" A_Index " is:`n" A_LoopField "`n`nContinue?")
        ; IfMsgBox, No
        ;     return
    }
}

; add info about the application to task bar
A_TrayMenu.Add("App Info", AppInfoMenu)


AppInfoMenu(*) {
    appInfoGui := Gui("+AlwaysOnTop", "AHK Text Expander Info")
    appInfoGui.OnEvent("Close", AppInfoGui_Close)
    appInfoGui.SetFont("s18", "Verdana")
    appInfoGui.AddText(, "AutoHotkey Text Expander")
    appInfoGui.SetFont("s10", "Verdana")
    appInfoGui.AddText("w500 h200", "This text expander allows you to automatically convert short phrases into long blocks of text. New shortcuts can be added in the included " hotstringFilename " file.`n`nExample, typing <ate will expand into `"AutoHotkey Text Expander`"`n`nBuilt In Hotstring:`n<now = DateTime(MM/dd/yyyy hh:mm:ss)`n`nIf you would like to keep track of how often your hotstrings are used, create a " hotstringCounterFilename " in the root folder where the application is stored and it will keep a running tally.")
    appInfoGui.Show()
}

AppInfoGui_Close(*) {
    ; gui automatically closes when this event handler returns
}

executeHotstring(HotStringExtendedText, hotstringCounterObject, *) {

    ; check if extended text has <<input>> at the beginning, get input from user and replace all instances of <<template>>
    if(SubStr(HotStringExtendedText, 1, 9) == "<<input>>"){
        Sleep(50)
        UserInputValue := createTextInputWindowAndWait()
        ; remove the <<input>> prefix first
        HotStringExtendedText := SubStr(HotStringExtendedText, 10)
        ; replace all <<template>> placeholders with user input
        HotStringExtendedText := StrReplace(HotStringExtendedText, "<<template>>", UserInputValue)
    }
    
    ; send hotstring ExtendedText
    SendInput(HotStringExtendedText)

    ; update hotstring counter object
    thisHotkeyValue := StrReplace(A_ThisHotkey, ":*:")
    if (hotstringCounterObject.HasOwnProp(thisHotkeyValue)){
        hotstringCounterObject.%thisHotkeyValue% += 1
    } else {
        hotstringCounterObject.%thisHotkeyValue% := 1
    }
}

retrieveObjectFromFile(filename) {
    tempObject := {}
    if FileExist(filename) {
        fileContent := FileRead(filename)
        lines := StrSplit(fileContent, "`n")
        for index, line in lines {
            if (line != "") {
                lineData := StrSplit(line, ",")
                if (lineData.Length >= 2) {
                    tempObject.%lineData[1]% := lineData[2]
                }
            }
        }
    }
    return tempObject
}

saveObjectRowsToTextFile(targetObject, filepath) {
    if (targetObject.OwnProps().Length = 0){
        ; MsgBox("DEBUG: No Data To Save")
        return
    }
    
    try {
        FileDelete(filepath)
    }
    
    combinedData := ""
    for key, value in targetObject.OwnProps() {
        combinedData .= key "," value "`n"
    }
    FileAppend(combinedData, filepath)
    MsgBox("DEBUG: Data Saved To File`n`n" combinedData)
}

createTextInputWindowAndWait() {
    inputGui := Gui("+AlwaysOnTop", "Input Text")
    inputGui.SetFont("s14", "Verdana")
    editControl := inputGui.AddEdit("vUserInputValue x12 y10 w326")
    submitBtn := inputGui.AddButton("Default x12 y40 w100 h30", "Submit")
    
    userInput := ""
    guiClosed := false
    
    ; define event handler functions
    submitBtn.OnEvent("Click", SubmitClick)
    inputGui.OnEvent("Close", GuiClose)
    
    inputGui.Show("w350 h80")
    
    ; wait for the window to be closed
    while (!guiClosed) {
        Sleep(50)
    }
    
    inputGui.Destroy()
    return userInput
    
    ; local event handler functions
    SubmitClick(*) {
        userInput := editControl.Text
        guiClosed := true
        inputGui.Destroy()
    }
    
    GuiClose(*) {
        if (!guiClosed) {
            userInput := editControl.Text
        }
        guiClosed := true
    }
}

; built in hotstrings
Hotstring(":*:<now", NowHotstring)

NowHotstring(*) {
    CurrentDateTime := FormatTime(, "MM/dd/yyyy hh:mm:ss")
    SendInput(CurrentDateTime)
}