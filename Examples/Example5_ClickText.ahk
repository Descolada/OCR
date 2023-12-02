#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

Run "https://www.w3schools.com/tags/att_input_type_checkbox.asp"
WinWaitActive "HTML input type",,10
if !WinActive("HTML input type") {
    MsgBox "Failed to find test window!"
    ExitApp
}

; Wait for text "Yourself" to appear, case-insensitive search, indefinite wait. Search only the active window.
result := OCR.WaitText("Yourself",, OCR.FromWindow.Bind(OCR, "A"))
; Find the Word for "Yourself" in the result, and click it.
result.Click(result.FindString("Yourself"))
; Wait for text to appear, that matches RegExMatch with needle "I have a bike(\s|$)". 
; RegEx matching is used here to accept either a space at the end or the end of string, because
; it might be in the middle of the found text or at the end.
; Search only the active window.
result := OCR.WaitText("I have a bike(\s|$)",, OCR.FromWindow.Bind(OCR,"A"),,RegExMatch)
; Here we don't have to use RegEx, because the string will be split by spaces and compared word-by-word.
result.Click(result.FindString("I have a bike"))