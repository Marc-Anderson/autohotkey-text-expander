#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
; VERSION 112
; This text expander is a work in progress

; name of file containing hotstrings
hotstringFilename := "hotstrings.xlsx"
; name of worksheet containing the hotstrings
hotstringWorksheet := "Templates"

; initialize the XL variable
XL :=

; fetch the correct file path
tgtFilePath := A_ScriptDir "\" hotstringFilename

MsgBox, 1, AutoHotkey Text Expander, This text expander allows you to automatically convert short phrases into long blocks of text. New shortcuts can be added in the included hotstrings.xlsx file.`n`nExample, typing <ate will expand into "AutoHotkey Text Expander"`n`nBuilt In Hotstring:`n<now = DateTime(MM/dd/yyyy hh:mm:ss)
IfMsgBox, Cancel 
    ExitApp
IfMsgBox, OK
    try {
        ; Check if excel is active
        XL := ComObjectActive("Excel.Application")
    } catch {
        ; If Excel is not active, create an instance
        XL := ComObjCreate("Excel.Application")
    }
    ; MsgBox, % IsObject(XL) ; Is excel an object?
    try {
        ; Check if the workbook exists
        XL.Workbooks.Open(tgtFilePath)
    } catch {
        ; Quit the application
        if(XL.Workbooks.Count = 1){
            XL.Quit
        }
        ; Alert the user that the file was not found
        MsgBox, The necessary workbook was not found. Please create a hotstrings.xlsx file to continue.
        ExitApp
        Sleep, 2000
    }

    ; Make Excel Visible
    XL.Visible := 1

    ; select the sheet name containing the templates
    tgtSheet := XL.Worksheets(hotstringWorksheet)

    ; activate the sheet with hotstrings
    tgtSheet.Activate

    ; Sort the data by column C so it can loop over all of the hotstrings without empty cells
    tgtSheet.UsedRange.Offset(1).Sort(XL.Columns(3), 1)

    ; Loop through all of the active cells
    while(tgtSheet.Range("C" . A_Index).Value != "") {

        ; Skip the first row as they are the header
        IfLess, A_Index, 2, Continue

        ; assign the value in column C of the current row to HotStringShortCut variable
        HotStringShortCut := tgtSheet.Range("C" . A_Index).Value

        ; assign the value in column D of the current row to HotStringExtended variable
        HotStringExtended := tgtSheet.Range("D" . A_Index).Value

        ; replaces any exclamation points and carriage returns with appropriate characters 
        HotStringExtended := StrReplace(StrReplace(StrReplace(HotStringExtended, "!","{!}"),"`r","{enter}"),A_Space A_Space,"{space 2}")

        ; assign both variables to a hotkey
        hotstring(":*:" HotStringShortCut, HotStringExtended)
        ; MsgBox, 4, , %HotStringShortCut% - %HotStringExtended%`n`nContinue?

    }
    ; tell excel it's save so it wont harass you and close the document
    XL.Application.ActiveWorkbook.saved := true
    wbCount := XL.Workbooks.Count    
    if(XL.Workbooks.Count = 1){
        XL.Quit
    } else {
        for WB in XL.Workbooks {
            wbName := WB.Name
            if(WB.Name = hotstringFilename){
                WB.Close
            }
        }
    }

    :*:<now::
        FormatTime, CurrentDateTime,, MM/dd/yyyy hh:mm:ss 
        SendInput, %CurrentDateTime%
        return