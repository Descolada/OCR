#include ..\Lib\OCR.ahk

CoordMode "Mouse", "Window"
Loop {
    ib := InputBox("Insert search phrase to find from active window: ", "OCR")
    if ib.Result != "OK"
        ExitApp
    result := OCR.FromWindow("A",,2)
    for word in result.Words {
        if word.Text = ib.Value {
            loc := word.Location
            ; MouseMove is set to CoordMode Window, so no conversion necessary
            MouseMove loc.x, loc.y
            ; Coordinates are relative to the window, so convert them into screen coordinates
            WinGetPos(&X, &Y,,,"A")
            Highlight(loc.x+X, loc.y+Y, loc.w, loc.h, 3000)
            break
        }
    }
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