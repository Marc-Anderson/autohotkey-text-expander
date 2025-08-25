#Requires AutoHotkey >=2.0
; #NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode("Input")  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir(A_ScriptDir)  ; Ensures a consistent starting directory.
SetTitleMatchMode(2)

; name and path of file containing hotstrings
hotstringFilename := "hotstrings.xlsx"
hotstringFilepath := A_ScriptDir "\" hotstringFilename

; name of worksheet containing the hotstrings
hotstringWorksheetName := "Templates"

; name and path of file for keeping track of use for each hotstring
hotstringCounterFilename := "hotstring-counter.txt"
hotstringCounterFilepath := A_WorkingDir "\" hotstringCounterFilename

; name and path of file for splash image
splashImageFilename := "splashfile300x100.png"
splashImageFilepath := A_ScriptDir "\" splashImageFilename

; initialize the XL variable
XL := ""

; show error if the hotstring file doesnt exist
if !FileExist(hotstringFilepath) {
    
    ; Alert the user that the file was not found
    MsgBox("The necessary workbook was not found. Please create a " hotstringFilename " file to continue.")
    ExitApp()
    Sleep(2000)
}

; create splash image for 
splashGui := Gui("+AlwaysOnTop -Caption", "Splash")
splashGui.SetFont("s14", "Verdana")
splashGui.AddText("x0 y50 w330 Center", "AutoHotkey Text Expander")
splashGui.SetFont("s8", "Verdana")
if FileExist(splashImageFilepath){
    splashGui.AddPicture("x15 y15 w300 h100", splashImageFilepath)
}
splashGui.AddText("x275 y115 Center", "Loading...")
splashGui.Show("NoActivate w330 h130")
WinSetTransparent(150, splashGui)

; try to load the hotstring file
try {
    ; Check if excel is active
    XL := ComObjActive("Excel.Application")
} catch {
    ; If Excel is not active, create an instance
    XL := ComObject("Excel.Application")
}
; MsgBox("Is excel an object? " . IsObject(XL))
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
    splashGui.Destroy()

    ; Alert the user that there was an error opening the hotstring file
    MsgBox("Either the necessary workbook was not found or there was another error opening the hotstring file. Please create a " hotstringFilename " file and check your excel installation to continue.")
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

; select the sheet name containing the templates
hotstringWorksheet := XL.Worksheets(hotstringWorksheetName)

; activate the sheet with hotstrings
hotstringWorksheet.Activate

; Sort the data by column C so it can loop over all of the hotstrings without empty cells
hotstringWorksheet.UsedRange.Offset(1).Sort(XL.Columns(3), 1)

; Loop through all of the active cells
while(hotstringWorksheet.Range("C" . A_Index).Value != "") {

    ; Skip the first row as they are the header
    if (A_Index < 2) {
        continue
    }

    ; assign the value in column C of the current row to HotStringShortCut variable
    HotStringShortCut := hotstringWorksheet.Range("C" . A_Index).Value

    ; assign the value in column D of the current row to HotStringExtended variable
    HotStringExtended := hotstringWorksheet.Range("D" . A_Index).Value

    ; replaces any exclamation points and carriage returns with appropriate characters 
    HotStringExtended := StrReplace(StrReplace(StrReplace(HotStringExtended, "!","{!}"),"`r","{enter}"),A_Space A_Space,"{space 2}")

    ; assign both variables to a hotkey and execute via function
    Hotstring(":*:" HotStringShortCut, executeHotstring.Bind(HotStringExtended, hotstringCounterObject), "On")
    ; MsgBox("HotStringShortCut: " HotStringShortCut "`nHotStringExtended: " HotStringExtended "`n`nContinue?")

}

; tell excel it's save so it wont harass you and close the document
XL.Application.Workbooks(hotstringFilename).saved := true
XL.Application.Workbooks(hotstringFilename).Close()

if(XL.Workbooks.Count = 0){
    XL.Application.Quit()
}

; clear any unused variables
hotstringWorksheet := ""
HotStringExtended := ""
HotStringShortCut := ""
XL := ""

; add info about the application to task bar
A_TrayMenu.Add("App Info", AppInfoMenu)

; remove the splash screen
splashGui.Destroy()

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
    ; MsgBox("DEBUG: Data Saved To File`n`n" combinedData)
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