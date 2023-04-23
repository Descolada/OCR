#include ..\Lib\OCR.ahk

CoordMode "Mouse", "Window"
Loop {
    ib := InputBox("Insert search phrase to find from active window: ", "OCR")
    if ib.Result != "OK"
        ExitApp
    result := OCR.FromWindow("A",,2)
    try found := result.FindString(ib.Value)
    catch {
        MsgBox 'Phrase "' ib.Value '" not found!'
        continue
    }
    ; MouseMove is set to CoordMode Window, so no coordinate conversion necessary
    MouseMove found.x, found.y
    result.Highlight(found)
    break
}