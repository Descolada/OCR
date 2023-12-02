#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

CoordMode "Mouse", "Screen"
Loop {
    ib := InputBox("Insert RegEx search phrase to find all matches from Desktop: ", "OCR")
    Sleep 500 ; Small delay to wait for the InputBox to close
    if ib.Result != "OK"
        ExitApp
    result := OCR.FromDesktop(,2)
    found := result.FindStrings(ib.Value,,RegExMatch)
    if !found.Length {
        MsgBox 'Phrase "' ib.Value '" not found!'
        continue
    }
    for match in found {
        ; MouseMove is set to CoordMode Screen, so no coordinate conversion necessary
        MouseMove match.x, match.y
        result.Highlight(match)
    } 
    break
}