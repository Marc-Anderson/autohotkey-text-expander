#NoEnv  					; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  						; Enable warnings to assist with detecting common errors.
SendMode Play  					; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  			; Set starting directory to the directory with the script.
#SingleInstance					;Prevent more than one of the same script from running concurrently.
#persistent					;Keep script active with timers running.
SetTitleMatchMode,2				;Easily find window titles.


MsgBox, 1, AutoHotkey Text Expander, This text expander allows you to automatically convert short phrases into long blocks of text. New shortcuts can be added in the included hotstrings.csv file.`n`nExample, typing <ate will expand into "AutoHotkey Text Expander"`n`nBuilt In Hotstring:`n<now = DateTime(MM/dd/yyyy hh:mm:ss)
IfMsgBox, Cancel 
    ExitApp
IfMsgBox, OK
    ; loop through the csv file one line at a time
    Loop, read, %A_ScriptDir%\hotstrings.csv
    {
        ; save the line number to a variable
        LineNumber := A_Index -1

        ; skip the first row of content
        IfEqual, A_Index, 1, Continue

        ; parse each cell into variables
        Loop, parse, A_LoopReadLine, CSV
            {
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

    :*:<now::
        FormatTime, CurrentDateTime,, MM/dd/yyyy hh:mm:ss 
        SendInput, %CurrentDateTime%
        return