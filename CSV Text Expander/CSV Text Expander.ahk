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

MsgBox, 1, AutoHotkey Text Expander, This text expander allows you to automatically convert short phrases into long blocks of text. New shortcuts can be added in the included %hotstringFilename% file.`n`nExample, typing <ate will expand into "AutoHotkey Text Expander"`n`nBuilt In Hotstring:`n<now = DateTime(MM/dd/yyyy hh:mm:ss)
IfMsgBox, Cancel 
    ExitApp
    ; Sleep, 2000

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
            hotstring(":*:" HotStringShortCut, HotStringExtended)
        }

        ; %LineNumber% is the current row
        ; %A_Index% is the current column
        ; %A_LoopField% is the current field
        
        ; MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
        ; IfMsgBox, No
        ;     return
    }
}

; built in hotstrings
:*:<now::
    FormatTime, CurrentDateTime,, MM/dd/yyyy hh:mm:ss 
    SendInput, %CurrentDateTime%
    return