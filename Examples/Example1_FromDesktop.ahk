#include ..\Lib\OCR.ahk

result := OCR.FromDesktop()
MsgBox "All text from desktop: `n" result.Text

MsgBox "Press OK to start highlighting all found lines.`n(This might take a while)"
for line in result.Lines
    result.Highlight(line)
ExitApp