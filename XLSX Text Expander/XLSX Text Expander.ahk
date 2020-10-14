#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
; VERSION 100

; This app works but is incomplete. still needs referential file location instead of typing it into the file. 

; initialize the XL variable
XL :=

; fetch the correct file path
tgtFilePath := A_ScriptDir "\hotstrings.xlsx"


MsgBox, 1, AutoHotkey Text Expander, This text expander app allows you to automatically convert short phrases into long blocks of text.`n`nFor example, typing <now will automatically expand into the current date and time. All shortcuts used in this app are prefixed with the < symbol.`n`nNew shortcuts can be added in the hotstrings.xlsx file included with this app.
IfMsgBox, Cancel 
    Return
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
    XL.Visible := 0

    ; select the sheet name containing the templates
    tgtSheet := XL.Worksheets("Templates")

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
        hotstring(":*:" "<" HotStringShortCut, HotStringExtended)
        ; MsgBox, 4, , %HotStringShortCut% - %HotStringExtended%`n`nContinue?

    }
    ; tell excel it's save so it wont harass you and close the document
    XL.Application.ActiveWorkbook.saved := true
    if(XL.Workbooks.Count = 1){
        XL.Quit
    } else {
        XL.Application.ActiveWorkbook.Close
    }

    :*:<now::
        FormatTime, CurrentDateTime,, MM/dd/yyyy hh:mm:ss 
        SendInput, %CurrentDateTime%
        return