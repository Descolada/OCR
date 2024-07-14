#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

WinTitle := "ahk_exe notepad.exe"
; Run Notepad and fill with dummy text or activate a pre-existing window
if !(hWnd := WinExist(WinTitle)) {
    Run "notepad.exe"
    WinWaitActive WinTitle
    hWnd := WinExist("A")
    Send "Lorem ipsum "
    Sleep 40
} else {
    WinActivate hWnd
    WinWaitActive hWnd
}
; Send Notepad to the background for demonstration purposes
WinMoveBottom(hWnd)
Sleep 60

WinGetClientPos(&X, &Y, &W, &H, hWnd)

; OCR only the top left quarter of the client area
result := OCR.FromWindow(hWnd,, 2, {X:0, Y:0, W:W//2, H:H//2, onlyClientArea:1})
MsgBox "Found in client area X:0 Y:0 W:" W//2 " H:" H//2 ":`n" result.Text
for line in result.Lines ; Highlight all lines for 2 seconds
    result.Highlight(line, -2000)
Sleep 2000

; OCR only the bottom right quarter of the client area
result := OCR.FromWindow(hWnd,, 2, {X:W//2, Y:H//2, W:W, H:H, onlyClientArea:1})
MsgBox "Found in client area X:" W//2 " Y:" H//2 " W:" W " H:" H ":`n" result.Text
for line in result.Lines ; Highlight all lines for 2 seconds
    result.Highlight(line, -2000)
Sleep 2000

ExitApp