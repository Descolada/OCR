#include ..\Lib\OCR.ahk

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"
Send "Lorem ipsum "
Sleep 40

for word in OCR.FromWindow("A",,2).Words
    if word.Text = "Lorem" {
        found := word
        break
    }
if !IsSet(found) {
    MsgBox '"ipsum" was not found in Notepad!'
    ExitApp
}

CoordMode "Mouse", "Window"
loc := found.Location
MouseClickDrag("Left", loc.x, loc.y, loc.x + loc.w, loc.y + loc.h)