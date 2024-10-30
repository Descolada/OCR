#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

CoordMode "Mouse", "Screen"
CoordMode "ToolTip", "Screen"

DllCall("SetThreadDpiAwarenessContext", "ptr", -3) ; Needed for multi-monitor setups with differing DPIs
; OCR.PerformanceMode := 1 ; Uncommenting this makes the OCR more performant, but also more CPU-heavy

global w := 150, h := 50, minsize := 5, step := 3
Loop {
    MouseGetPos(&x, &y)
    Highlight(x-w//2, y-h//2, w, h)
    ToolTip(g_CurrentText := OCR.FromRect(x-w//2, y-h//2, w, h,,2).Text, , y+h//2+10)
}

Right::global w+=step
Left::global w-=(w < minsize ? 0 : step)
Up::global h+=step
Down::global h-=(h < minsize ? 0 : step)

; Ctrl+Shift+c to copy the current OCR text to clipboard
^+c::A_Clipboard := g_CurrentText, ToolTip("OCR text copied to clipboard!"), Sleep(2000), ToolTip()

Highlight(x?, y?, w?, h?, showTime:=0, color:="Red", d:=2) {
	static guis := []

	if !IsSet(x) {
        for _, r in guis
            r.Destroy()
        guis := []
		return
    }
    if !guis.Length {
        Loop 4
            guis.Push(Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000000"))
    }
	Loop 4 {
		i:=A_Index
		, x1:=(i=2 ? x+w : x-d)
		, y1:=(i=3 ? y+h : y-d)
		, w1:=(i=1 or i=3 ? w+2*d : d)
		, h1:=(i=2 or i=4 ? h+2*d : d)
		guis[i].BackColor := color
		guis[i].Show("NA x" . x1 . " y" . y1 . " w" . w1 . " h" . h1)
	}
	if showTime > 0 {
		Sleep(showTime)
		Highlight()
	} else if showTime < 0
		SetTimer(Highlight, -Abs(showTime))
}