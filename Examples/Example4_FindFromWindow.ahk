#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

Loop {
    ib := InputBox("Insert search phrase to find from active window: ", "OCR")
    Sleep 100 ; Small delay to wait for the InputBox to close
    if ib.Result != "OK"
        ExitApp
    result := OCR.FromWindow("A", {scale:2})
    try found := result.FindString(ib.Value)
    catch {
        MsgBox 'Phrase "' ib.Value '" not found!'
        continue
    }
    ; MouseMove is set to CoordMode Window, so no coordinate conversion necessary
    MouseMove found.x, found.y
    found.Highlight()
    break
}