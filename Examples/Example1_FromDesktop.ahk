#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

result := OCR.FromDesktop()
MsgBox "All text from desktop: `n" result.Text

MsgBox "Press OK to highlight all found lines for 3 seconds."
for line in result.Lines
    result.Highlight(line, -3000)