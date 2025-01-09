#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"
Send "Lorem ipsum "
Sleep 40

result := OCR.FromWindow("A", {scale:2})
try found := result.FindString("Lorem")
if !IsSet(found) {
    MsgBox '"Lorem" was not found in Notepad!'
    ExitApp
}

found.Highlight()

CoordMode "Mouse", "Window"
MouseClickDrag("Left", found.x, found.y, found.x + found.w, found.y + found.h)