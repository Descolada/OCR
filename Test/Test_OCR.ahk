#include "..\Lib\OCR.ahk"
#Include "DUnit.ahk"
DUnit("C", OCRTestSuite)

class OCRTestSuite {
    static Fail() {
        throw Error()
    }
    static __FindWord(result, needle:="Lorem") {
        for word in result.Words
            if word.Text = needle
                return word.BoundingRect
        throw Error("Lorem not found", -1)
    }
    static __CompDiff(loc1, loc2, diff:=4) {
        return Abs(loc1.x-loc2.x + loc1.y-loc2.y + loc1.w-loc2.w + loc1.h-loc2.h) < diff
    }
    Begin() {
        Run "notepad.exe"
        WinWaitActive "ahk_exe notepad.exe"
        WinMove(0,0,1530,876)
        ControlSetText("Lorem ipsum ", "Edit1")
        ControlSend("{End}", "Edit1")
    }
    End() {
        WinClose "ahk_exe notepad.exe"
    }
    Test_FromDesktop() {
        result := OCR.FromDesktop()
        DUnit.Assert(InStr(result.Text, "Lorem ipsum"), '"Lorem ipsum" missing')
        DUnit.Assert(InStr(result.Text, "Type here to search"), '"Type here to search" missing')
        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:15, w:58, x:20, y:82}))
    }
    Test_FromWindow() {
        result := OCR.FromWindow("A") ; window, mode=2
        DUnit.Assert(InStr(result.Text, "Lorem ipsum") '"Lorem ipsum" missing')
        DUnit.Assert(InStr(result.Text, "File Edit Format View Help"), "Menubar missing")
        DUnit.Assert(!InStr(result.Text, "Type here to search"), '"Type here to search" NOT missing')

        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:14, w:57, x:20, y:81}))

        result := OCR.FromWindow("A",,,1) ; client, mode=2
        DUnit.Assert(InStr(result.Text, "Lorem ipsum"), '"Lorem ipsum" missing')
        DUnit.Assert(!InStr(result.Text, "File Edit Format View Help"), "Menubar NOT missing")

        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:14, w:58, x:8, y:6}))

        result := OCR.FromWindow("A",,,,0) ; window, mode=0
        DUnit.Assert(InStr(result.Text, "Lorem ipsum"), '"Lorem ipsum" missing')
        DUnit.Assert(InStr(result.Text, "File Edit Format View Help"), "Menubar missing")

        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:14, w:57, x:20, y:81}))

        result := OCR.FromWindow("A",,,1,0) ; client, mode=2
        DUnit.Assert(InStr(result.Text, "Lorem ipsum"), '"Lorem ipsum" missing')
        DUnit.Assert(!InStr(result.Text, "File Edit Format View Help"), "Menubar NOT missing")

        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:14, w:58, x:8, y:6}))
    }
    Test_FromFile() {
        result := OCR.FromFile("Notepad.png")
        DUnit.Assert(InStr(result.Text, "Lorem ipsum dolor"), '"Lorem ipsum" missing')

        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:14, w:58, x:9, y:81}))
    }
    Test_GetAvailableLanguages() {
        out := OCR.GetAvailableLanguages()
        DUnit.Assert(InStr(out, "en-US"))
    }
    Test_Scaling() {
        result := OCR.FromDesktop(,2)
        DUnit.Assert(InStr(result.Text, "Lorem ipsum") '"Lorem ipsum" missing')
        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:15, w:58, x:22, y:82}, 8), "Got " DUnit.Print(OCRTestSuite.__FindWord(result)))
        result := OCR.FromDesktop(,3)

        DUnit.Assert(InStr(result.Text, "Lorem ipsum") '"Lorem ipsum" missing')
        DUnit.Assert(OCRTestSuite.__CompDiff(OCRTestSuite.__FindWord(result), {h:15, w:58, x:22, y:82}, 8), "Got " DUnit.Print(OCRTestSuite.__FindWord(result)))
    }
    Test_Filter() {
        result := OCR.FromWindow("A")
        DUnit.Equal(result.Filter((word) => word.Text = "Lorem").Text, "Lorem")
        DUnit.Assert(!InStr(result.Crop(500, 500, 1000, 1000).Text, "Lorem"))
    }
}