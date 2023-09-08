#include ..\Lib\OCR.ahk

CoordMode "Mouse", "Screen"
CoordMode "ToolTip", "Screen"

DllCall("SetThreadDpiAwarenessContext", "ptr", -3) ; Needed for multi-monitor setups with differing DPIs

Loop {
    MouseGetPos(&X, &Y)
    Highlight(x-75, y-25, 150, 50)
    ToolTip(OCR.FromRect(X-75, Y-25, 150, 50,,2).Text, , Y+40)
}

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