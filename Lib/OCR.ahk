#Requires AutoHotkey v2

/**
 * OCR library: a wrapper for the the UWP Windows.Media.Ocr library.
 * Based on the UWP OCR function for AHK v1 by malcev.
 * 
 * Ways of initiating OCR:
 * OCR(IRandomAccessStream, lang?)
 * OCR.FromDesktop(lang?, scale:=1)
 * OCR.FromRect(X, Y, W, H, lang?, scale:=1)
 * OCR.FromWindow(WinTitle?, lang?, scale:=1, onlyClientArea:=0, mode:=2)
 * OCR.FromFile(FileName, lang?)
 * OCR.FromBitmap(bitmap, lang?, scale:=1)
 * 
 * Additional methods:
 * OCR.GetAvailableLanguages()
 * OCR.LoadLanguage(lang:="FirstFromAvailableLanguages")
 * OCR.WaitText(needle, timeout:=-1, func?, casesense:=False, comparefunc?)
 *      Calls a func (the provided OCR method) until a string is found
 * OCR.WordsBoundingRect(words*)
 *      Returns the bounding rectangle for multiple words
 * OCR.ClearAllHighlights()
 *      Removes all highlights created by Result.Highlight
 * OCR.Cluster(objs, eps_x:=-1, eps_y:=-1, minPts:=1, compareFunc?, &noise?)
 *      Clusters objects (by default based on distance from eachother). Can be used to create more
 *      accurate "Line" results.
 * OCR.SortArray(arr, optionsOrCallback:="N", key?)
 *      Sorts an array in-place, optionally by object keys or using a callback function.
 * OCR.ReverseArray(arr)
 *      Reverses an array in-place.
 * OCR.UniqueArray(arr)
 *      Returns an array with unique values.
 * OCR.FlattenArray(arr)
 *      Returns a one-dimensional array from a multi-dimensional array
 * 
 * 
 * Properties:
 * OCR.MaxImageDimension
 * MinImageDimension is not documented, but appears to be 40 pixels (source: user FanaticGuru in AutoHotkey forums)
 * 
 * OCR returns an OCR results object:
 * Result.Text         => All recognized text
 * Result.TextAngle    => Clockwise rotation of the recognized text 
 * Result.Lines        => Array of all Line objects
 * Result.Words        => Array of all Word objects
 * Result.ImageWidth   => Used image width
 * Result.ImageHeight  => Used image height
 * 
 * Result.FindString(needle, i:=1, casesense:=False, wordCompareFunc?, searchArea?)
 *      Finds a string in the result
 * Result.FindStrings(needle, casesense:=False, wordCompareFunc?, searchArea?)
 *      Finds all strings in the result
 * Result.Click(Obj, WhichButton?, ClickCount?, DownOrUp?)
 *      Clicks an object (Word, FindString result etc)
 * Result.ControlClick(obj, WinTitle?, WinText?, WhichButton?, ClickCount?, Options?, ExcludeTitle?, ExcludeText?)
 *      ControlClicks an object (Word, FindString result etc)
 * Result.Highlight(obj?, showTime?, color:="Red", d:=2)
 *      Highlights a Word, Line, or object with {x,y,w,h} properties on the screen (default: 2 seconds), or removes the highlighting
 * Result.Filter(callback)
 *      Returns a filtered result object that contains only words that satisfy the callback function
 * Result.Crop(x1, y1, x2, y2)
 *      Crops the result object to contain only results from an area defined by points (x1,y1) and (x2,y2).
 * 
 * 
 * Line object:
 * Line.Text         => Recognized text of the line
 * Line.Words        => Array of Word objects for the Line
 * 
 * Word object:
 * Word.Text         => Recognized text of the word
 * Word.x,y,w,h      => Size and location of the Word. Coordinates are relative to the original image.
 * Word.BoundingRect => Bounding rectangle of the Word in format {x,y,w,h}. Coordinates are relative to the original image.
 * 
 * Additional notes:
 * Languages are recognized in BCP-47 language tags. Eg. OCR.FromFile("myfile.bmp", "en-AU")
 * Languages can be installed for example with PowerShell (run as admin): Install-Language <language-tag>
 *      or from Language settings in Settings.
 * Not all language packs support OCR though. A list of supported language can be gotten from 
 * Powershell (run as admin) with the following command: Get-WindowsCapability -Online | Where-Object { $_.Name -Like 'Language.OCR*' } 
 */
class OCR {
    static IID_IRandomAccessStream := "{905A0FE1-BC53-11DF-8C49-001E4FC686DA}"
         , IID_IPicture            := "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
         , IID_IAsyncInfo          := "{00000036-0000-0000-C000-000000000046}"

    class IBase {
        __New(ptr?) {
            if IsSet(ptr) && !ptr
                throw ValueError('Invalid IUnknown interface pointer', -2, this.__Class)
            this.DefineProp("ptr", {Value:ptr ?? 0})
        }
        __Delete() => this.ptr ? ObjRelease(this.ptr) : 0
    }

    static __New() {
        this.prototype.__OCR := this
        this.IBase.prototype.__OCR := this
        this.OCRLine.base := this.IBase, this.OCRLine.prototype.base := this.IBase.prototype ; OCRLine extends OCR.IBase
        this.OCRWord.base := this.IBase, this.OCRWord.prototype.base := this.IBase.prototype ; OCRWord extends OCR.IBase
        this.LanguageFactory := this.CreateClass("Windows.Globalization.Language", ILanguageFactory := "{9B0252AC-0C27-44F8-B792-9793FB66C63E}")
        this.BitmapTransform := this.CreateClass("Windows.Graphics.Imaging.BitmapTransform")
        this.BitmapDecoderStatics := this.CreateClass("Windows.Graphics.Imaging.BitmapDecoder", IBitmapDecoderStatics := "{438CCB26-BCEF-4E95-BAD6-23A822E58D01}")
        this.OcrEngineStatics := this.CreateClass("Windows.Media.Ocr.OcrEngine", IOcrEngineStatics := "{5BFFA85A-3384-3540-9940-699120D428A8}")
        ComCall(6, this.OcrEngineStatics, "uint*", &MaxImageDimension:=0)   ; MaxImageDimension
        this.MaxImageDimension := MaxImageDimension
        DllCall("Dwmapi\DwmIsCompositionEnabled", "Int*", &compositionEnabled:=0)
        this.CAPTUREBLT := compositionEnabled ? 0 : 0x40000000
    }

    /**
     * Returns an OCR results object for an IRandomAccessStream.
     * Images of other types should be first converted to this format (eg from file, from bitmap).
     * @param pIRandomAccessStream Pointer or an object containing a ptr to the stream
     * @param {String} lang OCR language. Default is first from available languages.
     * @returns {Ocr} 
     */
    __New(pIRandomAccessStream?, lang := "FirstFromAvailableLanguages") {
        if IsSet(lang) || !this.__OCR.HasOwnProp("CurrentLanguage")
            this.__OCR.LoadLanguage(lang?)
        ComCall(14, this.__OCR.BitmapDecoderStatics, "ptr", pIRandomAccessStream, "ptr*", BitmapDecoder:=this.__OCR.IBase())   ; CreateAsync
        this.__OCR.WaitForAsync(&BitmapDecoder)
        BitmapFrame := ComObjQuery(BitmapDecoder, IBitmapFrame := "{72A49A1C-8081-438D-91BC-94ECFC8185C6}")
        ComCall(12, BitmapFrame, "uint*", &width:=0)   ; get_PixelWidth
        ComCall(13, BitmapFrame, "uint*", &height:=0)   ; get_PixelHeight
        if (width > this.__OCR.MaxImageDimension) or (height > this.__OCR.MaxImageDimension)
           throw ValueError("Image is too big - " width "x" height ".`nIt should be maximum - " this.__OCR.MaxImageDimension " pixels")

        BitmapFrameWithSoftwareBitmap := ComObjQuery(BitmapDecoder, IBitmapFrameWithSoftwareBitmap := "{FE287C9A-420C-4963-87AD-691436E08383}")
        if width < 40 || height < 40 {
            scale := 40.0 / Min(width, height), this.ImageWidth := Ceil(width*scale), this.ImageHeight := Ceil(height*scale)
            ComCall(7, this.__OCR.BitmapTransform, "int", this.ImageWidth) ; put_ScaledWidth
            ComCall(9, this.__OCR.BitmapTransform, "int", this.ImageHeight) ; put_ScaledHeight
            ComCall(8, BitmapFrame, "uint*", &BitmapPixelFormat:=0) ; get_BitmapPixelFormat
            ComCall(9, BitmapFrame, "uint*", &BitmapAlphaMode:=0) ; get_BitmapAlphaMode
            ComCall(8, BitmapFrameWithSoftwareBitmap, "uint", BitmapPixelFormat, "uint", BitmapAlphaMode, "ptr", this.__OCR.BitmapTransform, "uint", IgnoreExifOrientation := 0, "uint", DoNotColorManage := 0, "ptr*", SoftwareBitmap:=this.__OCR.IBase()) ; GetSoftwareBitmapAsync
        } else {
            this.ImageWidth := width, this.ImageHeight := height
            ComCall(6, BitmapFrameWithSoftwareBitmap, "ptr*", SoftwareBitmap:=this.__OCR.IBase())   ; GetSoftwareBitmapAsync
        }
        this.__OCR.WaitForAsync(&SoftwareBitmap)

        ComCall(6, this.__OCR.OcrEngine, "ptr", SoftwareBitmap, "ptr*", OcrResult:=this.__OCR.IBase())   ; RecognizeAsync
        this.__OCR.WaitForAsync(&OcrResult)

        ; Cleanup
        this.__OCR.CloseIClosable(pIRandomAccessStream)
        this.__OCR.CloseIClosable(SoftwareBitmap)

        this.ptr := OcrResult.ptr, ObjAddRef(OcrResult.ptr)
    }
    __Delete() => this.ptr ? ObjRelease(this.ptr) : 0

    ; Gets the recognized text.
    Text {
        get {
            ComCall(8, this, "ptr*", &hAllText:=0)   ; get_Text
            buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hAllText, "uint*", &length:=0, "ptr")
            this.DefineProp("Text", {Value:StrGet(buf, "UTF-16")})
            this.__OCR.DeleteHString(hAllText)
            return this.Text
        }
    }

    ; Gets the clockwise rotation of the recognized text, in degrees, around the center of the image.
    TextAngle {
        get => (ComCall(7, this, "double*", &value:=0), value)
    }

    ; Returns all Line objects for the result.
    Lines {
        get {
            ComCall(6, this, "ptr*", LinesList:=this.__OCR.IBase()) ; get_Lines
            ComCall(7, LinesList, "int*", &count:=0) ; count
            lines := []
            loop count {
                ComCall(6, LinesList, "int", A_Index-1, "ptr*", OcrLine:=this.__OCR.OCRLine())               
                lines.Push(OcrLine)
            }
            this.DefineProp("Lines", {Value:lines})
            return lines
        }
    }

    ; Returns all Word objects for the result. Equivalent to looping over all the Lines and getting the Words.
    Words {
        get {
            local words := [], line, word
            for line in this.Lines
                for word in line.Words
                    words.Push(word)
            this.DefineProp("Words", {Value:words})
            return words
        }
    }

    /**
     * Clicks an object
     * @param Obj The object to click, which can be a OCR result object, Line, Word, or Object {x,y,w,h}
     * If this object (the one Click is called from) contains a "Relative" property (this is
     * added by default with OCR.FromWindow) containing a Hwnd property, then that window will be activated,
     * otherwise the Relative objects Window.xy/Client.xy properties values will be added to the x and y coordinates as offsets.
     */
    Click(Obj, WhichButton?, ClickCount?, DownOrUp?) {
        if !obj.HasProp("x") && InStr(Type(obj), "OCR")
            obj := this.__OCR.WordsBoundingRect(obj.Words)
        local x := obj.x, y := obj.y, w := obj.w, h := obj.h, mode := "Screen", hwnd
        if this.HasProp("Relative") {
            if this.Relative.HasOwnProp("Window")
                mode := "Window", hwnd := this.Relative.Window.Hwnd
            else if this.Relative.HasOwnProp("Client")
                mode := "Client", hwnd := this.Relative.Client.Hwnd
            if IsSet(hwnd) && !WinActive(hwnd) {
                WinActivate(hwnd)
                WinWaitActive(hwnd,,1)
            }
            x += this.Relative.%mode%.x, y += this.Relative.%mode%.y
        }
        oldCoordMode := A_CoordModeMouse
        CoordMode "Mouse", mode
        Click(x+w//2, y+h//2, WhichButton?, ClickCount?, DownOrUp?)
        CoordMode "Mouse", oldCoordMode
    }

    /**
     * ControlClicks an object
     * @param obj The object to click, which can be a OCR result object, Line, Word, or Object {x,y,w,h}
     * If the result object originates from OCR.FromWindow which captured only the client area,
     * then the result object will contain correct coordinates for the ControlClick. 
     * If OCR.FromWindow captured the Window area, then the Relative property
     * will contain Window property, and those coordinates will be adjusted to Client area.
     * Otherwise, if additionally a WinTitle is provided then the coordinates are treated as Screen 
     * coordinates and converted to Client coordinates.
     * @param WinTitle If WinTitle is set, then the coordinates stored in Obj will be converted to
     * client coordinates and ControlClicked.
     */
    ControlClick(obj, WinTitle?, WinText?, WhichButton?, ClickCount?, Options?, ExcludeTitle?, ExcludeText?) {
        if !obj.HasProp("x") && InStr(Type(obj), "OCR")
            obj := this.__OCR.WordsBoundingRect(obj.Words)
        local x := obj.x, y := obj.y, w := obj.w, h := obj.h, hWnd
        if this.HasProp("Relative") && (this.Relative.HasOwnProp("Client") || this.Relative.HasOwnProp("Window")) {
            mode := this.Relative.HasOwnProp("Client") ? "Client" : "Window"
            , obj := this.Relative.%mode%, x += obj.x, y += obj.y, hWnd := obj.hWnd
            if mode = "Window" {
                ; Window -> Client
                RECT := Buffer(16, 0), pt := Buffer(8, 0)
                DllCall("user32\GetWindowRect", "Ptr", hWnd, "Ptr", RECT)
                winX := NumGet(RECT, 0, "Int"), winY := NumGet(RECT, 4, "Int")
                NumPut("int", winX+x, "int", winY+y, pt)
                DllCall("user32\ScreenToClient", "Ptr", hWnd, "Ptr", pt)
                x := NumGet(pt,0,"int"), y := NumGet(pt,4,"int")
            }
        } else if IsSet(WinTitle) {
            hWnd := WinExist(WinTitle, WinText?, ExcludeTitle?, ExcludeText?)
            pt := Buffer(8), NumPut("int",x,pt), NumPut("int", y,pt,4)
            DllCall("ScreenToClient", "Int", Hwnd, "Ptr", pt)
            x := NumGet(pt,0,"int"), y := NumGet(pt,4,"int")
        } else
            throw TargetError("ControlClick needs to be called either after a OCR.FromWindow result or with a WinTitle argument")
            
        ControlClick("X" (x+w//2) " Y" (y+h//2), hWnd,, WhichButton?, ClickCount?, Options?)
    }

    /**
     * Highlights an object on the screen with a red box
     * @param obj The object to highlight. which can be a OCR result object, Line, Word, or Object {x,y,w,h}
     * If this object (the one Highlight is called from) contains a "Relative" property (this is
     * added by default with OCR.FromWindow), then its values will be added to the x and y coordinates as offsets.
     * @param {number} showTime Default is 2 seconds.
     * * Unset - if highlighting exists then removes the highlighting, otherwise pauses for 2 seconds
     * * 0 - Indefinite highlighting
     * * Positive integer (eg 2000) - will highlight and pause for the specified amount of time in ms
     * * Negative integer - will highlight for the specified amount of time in ms, but script execution will continue
     * * "clear" - removes the highlighting unconditionally
     * * "clearall" - remove highlightings from all OCR objects
     * @param {string} color The color of the highlighting. Default is red.
     * @param {number} d The border thickness of the highlighting in pixels. Default is 2.
     * @returns {OCR}
     */
    Highlight(obj?, showTime?, color:="Red", d:=2) {
        static Guis := Map()
        local x, y, w, h, key, resultObjs, key2, oObj, rect, ResultGuis, GuiObj, iw, ih
        ; obj set & showTime unset => either highlights for 2s, or removes highlight
        ; obj set & clear => removes highlight
        ; obj unset => clears all highlights unconditionally
        if IsSet(showTime) && showTime = "clearall" {
            for key, resultObjs in Guis { ; enum all OCR result objects
                for key2, oObj in resultObjs {
                    try oObj.GuiObj.Destroy()
                    SetTimer(oObj.TimerObj, 0)
                }
            }
            Guis := Map()
            return this
        }
        if !Guis.Has(this.ptr)
            Guis[this.ptr] := Map()

        if !IsSet(obj) {
            for key, oObj in Guis[this.ptr] { ; enumerate all previously used obj arguments and remove GUIs
                try oObj.GuiObj.Destroy()
                SetTimer(oObj.TimerObj, 0)
            }
            Guis.Delete(this.ptr)
            return this
        }
        ; Otherwise obj is set
        if !IsObject(obj)
            throw ValueError("First argument 'obj' must be an object", -1)
        ResultGuis := Guis[this.ptr]

        if (!IsSet(showTime) && ResultGuis.Has(obj)) || (IsSet(showTime) && showTime = "clear") {
                try ResultGuis[obj].GuiObj.Destroy()
                SetTimer(ResultGuis[obj].TimerObj, 0)
                ResultGuis.Delete(obj)
                return this
        } else if !IsSet(showTime)
            showTime := 2000

        if Type(obj) = this.__OCR.prototype.__Class ".OCRLine" || Type(obj) = this.__OCR.prototype.__Class
            rect := this.__OCR.WordsBoundingRect(obj.Words*)
        else 
            rect := obj
        x := rect.x, y := rect.y, w := rect.w, h := rect.h
        if this.HasProp("Relative") && this.Relative.HasOwnProp("Screen")
            x += this.Relative.Screen.X, y += this.Relative.Screen.Y

        if !ResultGuis.Has(obj) {
            ResultGuis[obj] := {}
            ResultGuis[obj].GuiObj := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000000")
            ResultGuis[obj].TimerObj := ObjBindMethod(this, "Highlight", obj, "clear")
        }
        GuiObj := ResultGuis[obj].GuiObj
        GuiObj.BackColor := color
        iw:= w+d, ih:= h+d, w:=w+d*2, h:=h+d*2, x:=x-d, y:=y-d
        WinSetRegion("0-0 " w "-0 " w "-" h " 0-" h " 0-0 " d "-" d " " iw "-" d " " iw "-" ih " " d "-" ih " " d "-" d, GuiObj.Hwnd)
        GuiObj.Show("NA x" . x . " y" . y . " w" . w . " h" . h)

        if showTime > 0 {
            Sleep(showTime)
            this.Highlight(obj)
        } else if showTime < 0
            SetTimer(ResultGuis[obj].TimerObj, -Abs(showTime))
        return this
    }
    ClearHighlight(obj) => this.Highlight(obj, "clear")
    static ClearAllHighlights() => this.Prototype.Highlight(,"clearall")

    /**
     * Finds a string in the search results. Returns {x,y,w,h,Words} where Words contains an array of the matching Word objects.
     * @param needle The string to find
     * @param {number} i Which occurrence of needle to find
     * @param {number} casesense Comparison case-sensitivity. Default is False/Off.
     * @param wordCompareFunc Optionally a custom word comparison function. Accepts two arguments,
     *     neither of which should contain spaces. 
     *     When using RegExMatch as wordCompareFunc note that a "space" will split the RegEx into multiple parts.
     *     Eg. "\w+   \d+" will actually match for a word satisfying "\w+" followed by a word satisfying "\d+"
     * @param searchArea Optionally a {x1,y1,x2,y2} object defining the search area inside the result object
     * @returns {Object} 
     */
    FindString(needle, i:=1, casesense:=False, wordCompareFunc?, searchArea?) {
        local line, counter, found, x1, y1, x2, y2, splitNeedle, result, word
        if !(needle is String)
            throw TypeError("Needle is required to be a string, not type " Type(needle), -1)
        if needle == ""
            throw ValueError("Needle cannot be an empty string", -1)
        splitNeedle := StrSplit(RegExReplace(needle, " +", " "), " "), needleLen := splitNeedle.Length
        if !IsSet(wordCompareFunc)
            wordCompareFunc := casesense ? ((arg1, arg2) => arg1 == arg2) : ((arg1, arg2) => arg1 = arg2)
        If IsSet(searchArea) {
            x1 := searchArea.HasOwnProp("x1") ? searchArea.x1 : -100000
            y1 := searchArea.HasOwnProp("y1") ? searchArea.y1 : -100000
            x2 := searchArea.HasOwnProp("x2") ? searchArea.x2 : 100000
            y2 := searchArea.HasOwnProp("y2") ? searchArea.y2 : 100000
        }
        for line in this.Lines {
            if IsSet(wordCompareFunc) || InStr(l := line.Text, needle, casesense) {
                counter := 0, found := []
                for word in line.Words {
                    If IsSet(searchArea) && (word.x < x1 || word.y < y1 || word.x+word.w > x2 || word.y+word.h > y2)
                        continue
                    t := word.Text, len := StrLen(t)
                    if wordCompareFunc(splitNeedle[found.Length+1], t) {
                        found.Push(word)
                        if found.Length == needleLen {
                            if ++counter == i {
                                result := this.__OCR.WordsBoundingRect(found*)
                                result.Words := found
                                return result
                            } else
                                found := []
                        }
                    } else
                        found := []
                }
            }
        }
        throw TargetError('The target string "' needle '" was not found', -1)
    }

    /**
     * Finds all strings matching the needle in the search results. Returns an array of {x,y,w,h,Words} objects
     * where Words contains an array of the matching Word objects.
     * @param needle The string to find. 
     * @param {number} casesense Comparison case-sensitivity. Default is False/Off.
     * @param wordCompareFunc Optionally a custom word comparison function. Accepts two arguments,
     *     neither of which should contain spaces. 
     *     When using RegExMatch as wordCompareFunc note that a "space" will split the RegEx into multiple parts.
     *     Eg. "\w+   \d+" will actually match for a word satisfying "\w+" followed by a word satisfying "\d+"
     * @param searchArea Optionally a {x1,y1,x2,y2} object defining the search area inside the result object
     * @returns {Array} 
     */
    FindStrings(needle, casesense:=False, wordCompareFunc?, searchArea?) {
        local line, counter, found, x1, y1, x2, y2, splitNeedle, result, word
        if !(needle is String)
            throw TypeError("Needle is required to be a string, not type " Type(needle), -1)
        if needle == ""
            throw ValueError("Needle cannot be an empty string", -1)
        splitNeedle := StrSplit(RegExReplace(needle, " +", " "), " "), needleLen := splitNeedle.Length
        if !IsSet(wordCompareFunc)
            wordCompareFunc := casesense ? ((arg1, arg2) => arg1 == arg2) : ((arg1, arg2) => arg1 = arg2)
        If IsSet(searchArea) {
            x1 := searchArea.HasOwnProp("x1") ? searchArea.x1 : -100000
            y1 := searchArea.HasOwnProp("y1") ? searchArea.y1 : -100000
            x2 := searchArea.HasOwnProp("x2") ? searchArea.x2 : 100000
            y2 := searchArea.HasOwnProp("y2") ? searchArea.y2 : 100000
        }
        results := []
        for line in this.Lines {
            if IsSet(wordCompareFunc) || InStr(l := line.Text, needle, casesense) {
                counter := 0, found := []
                for word in line.Words {
                    If IsSet(searchArea) && (word.x < x1 || word.y < y1 || word.x+word.w > x2 || word.y+word.h > y2)
                        continue
                    t := word.Text, len := StrLen(t)
                    if wordCompareFunc(t, splitNeedle[found.Length+1]) {
                        found.Push(word)
                        if found.Length == needleLen {
                            result := this.__OCR.WordsBoundingRect(found*)
                            result.Words := found
                            results.Push(result)
                            counter := 0, found := [], result := unset
                        }
                    } else
                        found := []
                }
            }
        }
        return results
    }

    /**
     * Filters out all the words that do not satisfy the callback function and returns a new OCR.Result object
     * @param {Object} callback The callback function that accepts a OCR.Word object.
     * If the callback returns 0 then the word is filtered out (rejected), otherwise is kept.
     * @returns {OCR}
     */
    Filter(callback) {
        if !HasMethod(callback)
            throw ValueError("Filter callback must be a function", -1)
        local result := this.Clone(), line, croppedLines := [], croppedText := "", croppedWords := [], lineText := "", word
        ObjAddRef(result.ptr)
        for line in result.Lines {
            croppedWords := [], lineText := ""
            for word in line.Words {
                if callback(word)
                    croppedWords.Push(word), lineText .= word.Text " "
            }
            if croppedWords.Length {
                line := {Text:Trim(lineText), Words:croppedWords}
                line.base.__Class := this.__OCR.prototype.__Class ".OCRLine"
                croppedLines.Push(line)
                croppedText .= lineText
            }
        }
        result.DefineProp("Lines", {Value:croppedLines})
        result.DefineProp("Text", {Value:Trim(croppedText)})
        result.DefineProp("Words", this.__OCR.Prototype.GetOwnPropDesc("Words"))
        return result
    }

    /**
     * Crops the result object to contain only results from an area defined by points (x1,y1) and (x2,y2).
     * Note that these coordinates are relative to the result object, not to the screen.
     * @param {Integer} x1 x coordinate of the top left corner of the search area
     * @param {Integer} y1 y coordinate of the top left corner of the search area
     * @param {Integer} x2 x coordinate of the bottom right corner of the search area
     * @param {Integer} y2 y coordinate of the bottom right corner of the search area
     * @returns {OCR}
     */
    Crop(x1:=-100000, y1:=-100000, x2:=100000, y2:=100000) => this.Filter((word) => word.x >= x1 && word.y >= y1 && (word.x+word.w) <= x2 && (word.y+word.h) <= y2)

    class OCRLine {
        ; Gets the recognized text for the line.
        Text {
            get {
                ComCall(7, this, "ptr*", &hText:=0)   ; get_Text
                buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length:=0, "ptr")
                text := StrGet(buf, "UTF-16")
                this.__OCR.DeleteHString(hText)
                this.DefineProp("Text", {Value:text})
                return text
            }
        }

        ; Gets the Word objects for the line
        Words {
            get {
                ComCall(6, this, "ptr*", WordsList:=this.__OCR.IBase())   ; get_Words
                ComCall(7, WordsList, "int*", &WordsCount:=0)   ; Words count
                words := []
                loop WordsCount {
                   ComCall(6, WordsList, "int", A_Index-1, "ptr*", OcrWord:=this.__OCR.OCRWord())
                   words.Push(OcrWord)
                }
                this.DefineProp("Words", {Value:words})
                return words
            }
        }

        BoundingRect {
            get => this.DefineProp("BoundingRect", {Value:this.__OCR.WordsBoundingRect(this.Words*)}).BoundingRect
        }
        x {
            get => this.BoundingRect.x
        } 
        y {
            get => this.BoundingRect.y
        }
        w {
            get => this.BoundingRect.w
        }
        h {
            get => this.BoundingRect.h
        }
    }

    class OCRWord {
        ; Gets the recognized text for the word
        Text {
            get {
                ComCall(7, this, "ptr*", &hText:=0)   ; get_Text
                buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length:=0, "ptr")
                text := StrGet(buf, "UTF-16")
                this.__OCR.DeleteHString(hText)
                this.DefineProp("Text", {Value:text})
                return text
            }
        }

        /**
         * Gets the bounding rectangle of the text in {x,y,w,h} format. 
         * The bounding rectangles coordinate system will be dependant on the image capture method.
         * For example, if the image was captured as a rectangle from the screen, then the coordinates
         * will be relative to the left top corner of the rectangle.
         */
        BoundingRect {
            get {
                ComCall(6, this, "ptr", RECT := Buffer(16, 0))   ; get_BoundingRect
                this.DefineProp("x", {Value:Integer(NumGet(RECT, 0, "float"))})
                , this.DefineProp("y", {Value:Integer(NumGet(RECT, 4, "float"))})
                , this.DefineProp("w", {Value:Integer(NumGet(RECT, 8, "float"))})
                , this.DefineProp("h", {Value:Integer(NumGet(RECT, 12, "float"))})
                return this.DefineProp("BoundingRect", {Value:{x:this.x, y:this.y, w:this.w, h:this.h}}).BoundingRect
            }
        }
        x {
            get => this.BoundingRect.x
        }
        y {
            get => this.BoundingRect.y
        }
        w {
            get => this.BoundingRect.w
        }
        h {
            get => this.BoundingRect.h
        }
    }

    /**
     * Returns an OCR results object for an image file. Locations of the words will be relative to
     * the top left corner of the image.
     * @param FileName Either full or relative (to A_ScriptDir) path to the file.
     * @param lang OCR language. Default is first from available languages.
     * @returns {Ocr} 
     */
    static FromFile(FileName, lang?) {
        if (SubStr(FileName, 2, 1) != ":")
            FileName := A_ScriptDir "\" FileName
         if !FileExist(FileName) or InStr(FileExist(FileName), "D")
            throw TargetError("File `"" FileName "`" doesn't exist", -1)
         GUID := this.CLSIDFromString(this.IID_IRandomAccessStream)
         DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", FileName, "uint", Read := 0, "ptr", GUID, "ptr*", IRandomAccessStream:=this.IBase())
         return this(IRandomAccessStream, lang?)
    }

    /**
     * Returns an OCR results object for a given window. Locations of the words will be relative to the
     * window or client area, so for interactions use CoordMode "Window" or "Client". If onlyClientArea
     * contained relative coordinates then Result coordinates will also be relative to the captured area.
     * In that case offsets for Window/Client area are stored in Result.Relative.Client.x and y or .Window.x and y.
     * Additionally, Result.Relative.Screen.x and y are also stored. 
     * @param WinTitle A window title or other criteria identifying the target window.
     * @param lang OCR language. Default is first from available languages.
     * @param scale The scaling factor to use.
     * @param {Number, Object} onlyClientArea Whether only the client area or the whole window should be OCR-d
     *     This can also be an object which must contain {X,Y,W,H} (relative coordinates from where to OCR) 
     *     and optionally onlyClientArea property (0 or 1, default is 0).
     * @param {Number} mode Different methods of capturing the window. 0 = uses GetDC with BitBlt, 2 = uses PrintWindow. 
     * Add 1 to make a transparent window totally opaque. 
     * @returns {Ocr} 
     */
    static FromWindow(WinTitle:="", lang?, scale:=1, onlyClientArea:=0, mode:=2) {
        local result, rect, X, Y, W, H, x2, y2, hBitMap, hwnd
        if !(hWnd := WinExist(WinTitle))
            throw TargetError("Target window not found", -1)
        if DllCall("IsIconic", "uptr", hwnd)
            DllCall("ShowWindow", "uptr", hwnd, "int", 4)
        if mode&1 {
            oldStyle := WinGetExStyle(hwnd), i := 0
            WinSetTransparent(255, hwnd)
            While (WinGetTransparent(hwnd) != 255 && ++i < 30)
                Sleep 100
        }
        if IsObject(onlyClientArea) {
            if !onlyClientArea.HasOwnProp("onlyClientArea") 
                onlyClientArea.onlyClientArea := 0
            X := onlyClientArea.X, Y := onlyClientArea.Y, W := onlyClientArea.W, H := onlyClientArea.H, flagOnlyClientArea := onlyClientArea.onlyClientArea
        } else
            X := 0, Y := 0, W := 0, H := 0, flagOnlyClientArea := onlyClientArea
        If flagOnlyClientArea = 1 {
            DllCall("GetClientRect", "ptr", hwnd, "ptr", rc:=Buffer(16))
            if !W
                W := NumGet(rc, 8, "int"), H := NumGet(rc, 12, "int")
            pt:=Buffer(8, 0), NumPut("int64", 0, pt)
            , DllCall("ClientToScreen", "Ptr", hwnd, "Ptr", pt)
            , X += NumGet(pt,"int"), Y += NumGet(pt,4,"int")
        } else {
            rect := Buffer(16, 0)
            , DllCall("GetWindowRect", "UPtr", hwnd, "Ptr", rect, "UInt")
            , X += NumGet(rect, 0, "Int"), Y += NumGet(rect, 4, "Int")
            if !W
                x2 := NumGet(rect, 8, "Int"), y2 := NumGet(rect, 12, "Int")
                , W := Abs(Max(X, X2) - Min(X, X2)), H := Abs(Max(Y, Y2) - Min(Y, Y2))
        }
        hBitMap := this.CreateBitmap(X, Y, W, H, hWnd, scale, onlyClientArea, mode)
        ;this.DisplayHBitmap(hBitMap)
        if mode&1
            WinSetExStyle(oldStyle, hwnd)
        result := this(this.HBitmapToRandomAccessStream(hBitMap), lang?)
        , result.Relative := {Screen:{X:X, Y:Y}}
        if IsObject(onlyClientArea)
            result.Relative.%(flagOnlyClientArea = 1 ? "Client" : "Window")% := {X:onlyClientArea.X, Y:onlyClientArea.Y, Hwnd:hWnd}
        else
            result.Relative.%(flagOnlyClientArea = 1 ? "Client" : "Window")% := {X:0, Y:0, Hwnd:hWnd}
        this.NormalizeCoordinates(result, scale)
        return result
    }

    /**
     * Returns an OCR results object for the whole desktop. Locations of the words will be relative to
     * the screen (CoordMode "Screen") in a single-monitor setup. If "monitor" argument is specified
     * then coordinates might be relative to the monitor, whereas relative offsets will be stored in
     * Result.Relative.Screen.x and y properties. 
     * @param lang OCR language. Default is first from available languages.
     * @param scale The scaling factor to use.
     * @param monitor The monitor from which to get the desktop area. Default is primary monitor.
     *   If screen scaling between monitors differs, then use DllCall("SetThreadDpiAwarenessContext", "ptr", -3)
     * @returns {Ocr} 
     */
    static FromDesktop(lang?, scale:=1, monitor?) {
        MonitorGet(monitor?, &Left, &Top, &Right, &Bottom)
        return this.FromRect(Left, Top, Right-Left, Bottom-Top, lang?, scale)
    }

    /**
     * Returns an OCR results object for a region of the screen. Locations of the words will be relative
     * to the top left corner of the rectangle. The return object will contain Relative.Screen.x and y properties
     * which are the original x and y that FromRect was called with.
     * @param x Screen x coordinate
     * @param y Screen y coordinate
     * @param w Region width. Maximum is OCR.MaxImageDimension; minimum is 40 pixels (source: user FanaticGuru in AutoHotkey forums), smaller images will be scaled to at least 40 pixels.
     * @param h Region height. Maximum is OCR.MaxImageDimension; minimum is 40 pixels, smaller images will be scaled accordingly.
     * @param lang OCR language. Default is first from available languages.
     * @param scale The scaling factor to use. Larger number (eg 2) might improve the accuracy
     *     of the OCR, at the cost of speed.
     * @returns {Ocr} 
     */
    static FromRect(x, y, w, h, lang?, scale:=1) {
        local hBitmap := this.CreateBitmap(X, Y, W, H,,scale)
        , result := this(this.HBitmapToRandomAccessStream(hBitmap), lang?)
        result.Relative := {Screen:{x:x, y:y}}
        return this.NormalizeCoordinates(result, scale)
    }

    /**
     * Returns an OCR results object from a bitmap. Locations of the words will be relative
     * to the top left corner of the bitmap.
     * @param bitmap A pointer to a GDIP Bitmap object, or HBITMAP, or an object with a ptr property
     *  set to one of the two.
     * @param lang OCR language. Default is first from available languages.
     * @returns {ocr} 
     */
    static FromBitmap(bitmap, lang?, scale:=1) {
        local result, pDC, hBitmap, hBM2, oBM, oBM2, pBitmapInfo := Buffer(32, 0)

        if !DllCall("GetObject", "ptr", bitmap, "int", pBitmapInfo.Size, "ptr", pBitmapInfo) {
            DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "UPtr", bitmap, "UPtr*", &hBitmap:=0, "Int", 0xffffffff)
            DllCall("GetObject", "ptr", hBitmap, "int", pBitmapInfo.Size, "ptr", pBitmapInfo)
        } else
            hBitmap := bitmap

        W := NumGet(pBitmapInfo, 4, "int"), H := NumGet(pBitmapInfo, 8, "int")

        if scale != 1 || (W && H && (W < 40 || H < 40)) {
            sW := Integer(W * scale), sH := Integer(H * scale)

            hDC := DllCall("CreateCompatibleDC", "Ptr", 0, "UPtr")
            oBM := DllCall("SelectObject", "Ptr", hDC, "Ptr", hBitmap)
            pDC := DllCall("CreateCompatibleDC", "Ptr", hDC, "UPtr")
            hBM2 := DllCall("CreateCompatibleBitmap", "Ptr", hDC, "Int", Max(40, sW), "Int", Max(40, sH), "UPtr")
            oBM2 := DllCall("SelectObject", "Ptr", pDC, "Ptr", hBM2)
            if sW < 40 || sH < 40 ; Fills the bitmap so it's at least 40x40, which seems to improve recognition
                DllCall("StretchBlt", "Ptr", pDC, "Int", 0, "Int", 0, "Int", Max(40,sW), "Int", Max(40,sH), "Ptr", hDC, "Int", 0, "Int", 0, "Int", 1, "Int", 1, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY. 
            DllCall("StretchBlt", "Ptr", pDC, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", hDC, "Int", 0, "Int", 0, "Int", W, "Int", H, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY
            DllCall("SelectObject", "Ptr", pDC, "Ptr", oBM2)
            DllCall("SelectObject", "Ptr", hDC, "Ptr", oBM)
            DllCall("DeleteDC", "Ptr", pDC)
            DllCall("DeleteDC", "Ptr", hDC)
            result := this(this.HBitmapToRandomAccessStream(hBM2), lang?)
            this.NormalizeCoordinates(result, scale)
            DllCall("DeleteObject", "UPtr", hBM2)
            return result
        } 
        return this(this.HBitmapToRandomAccessStream(hBitmap), lang?)
    } 

    /**
     * Returns all available languages as a string, where the languages are separated by newlines.
     * @returns {String} 
     */
    static GetAvailableLanguages() {
        ComCall(7, this.OcrEngineStatics, "ptr*", &LanguageList := 0)   ; AvailableRecognizerLanguages
        ComCall(7, LanguageList, "int*", &count := 0)   ; count
        Loop count {
            ComCall(6, LanguageList, "int", A_Index - 1, "ptr*", &Language := 0)   ; get_Item
            ComCall(6, Language, "ptr*", &hText := 0)
            buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length := 0, "ptr")
            text .= StrGet(buf, "UTF-16") "`n"
            this.DeleteHString(hText)
            ObjRelease(Language)
        }
        ObjRelease(LanguageList)
        return text
    }

    /**
     * Loads a new language which will be used with subsequent OCR calls.
     * @param {string} lang OCR language. Default is first from available languages.
     * @returns {void} 
     */
    static LoadLanguage(lang:="FirstFromAvailableLanguages") {
        local hString, Language, OcrEngine
        if this.HasOwnProp("CurrentLanguage") && this.HasOwnProp("OcrEngine") && this.CurrentLanguage = lang
            return
        if (lang = "FirstFromAvailableLanguages")
            ComCall(10, this.OcrEngineStatics, "ptr*", OcrEngine:=this.IBase())   ; TryCreateFromUserProfileLanguages
        else {
            hString := this.CreateHString(lang)
            , ComCall(6, this.LanguageFactory, "ptr", hString, "ptr*", Language:=this.IBase())   ; CreateLanguage
            , this.DeleteHString(hString)
            , ComCall(9, this.OcrEngineStatics, "ptr", Language, "ptr*", OcrEngine:=this.IBase())   ; TryCreateFromLanguage
        }
        if (OcrEngine.ptr = 0)
            Throw Error(lang = "FirstFromAvailableLanguages" ? "Failed to use FirstFromAvailableLanguages for OCR:`nmake sure the primary language pack has OCR capabilities installed.`n`nAlternatively try `"en-us`" as the language." : "Can not use language `"" lang "`" for OCR, please install language pack.")
        this.OcrEngine := OcrEngine, this.CurrentLanguage := lang
    }

    /**
     * Returns a bounding rectangle {x,y,w,h} for the provided Word objects
     * @param words Word object arguments (at least 1)
     * @returns {Object}
     */
    static WordsBoundingRect(words*) {
        if !words.Length
            throw ValueError("This function requires at least one argument")
        local X1 := 100000000, Y1 := 100000000, X2 := -100000000, Y2 := -100000000, word
        for word in words {
            X1 := Min(word.x, X1), Y1 := Min(word.y, Y1), X2 := Max(word.x+word.w, X2), Y2 := Max(word.y+word.h, Y2)
        }
        return {X:X1, Y:Y1, W:X2-X1, H:Y2-Y1, X2:X2, Y2:Y2}
    }
    
    /**
     * Waits text to appear on screen. If the method is successful, then Func's return value is returned.
     * Otherwise nothing is returned.
     * @param needle The searched text
     * @param {number} timeout Timeout in milliseconds. Less than 0 is indefinite wait (default)
     * @param func The function to be called for the OCR. Default is OCR.FromDesktop
     * @param casesense Text comparison case-sensitivity
     * @param comparefunc A custom string compare/search function, that accepts two arguments: haystack and needle.
     *      Default is InStr. If a custom function is used, then casesense is ignored.
     * @returns {OCR} 
     */
    static WaitText(needle, timeout:=-1, func?, casesense:=False, comparefunc?) {
        local endTime := A_TickCount+timeout, result
        if !IsSet(func)
            func := this.FromDesktop
        if !IsSet(comparefunc)
            comparefunc := InStr.Bind(,,casesense)
        While timeout > 0 ? (A_TickCount < endTime) : 1 {
            result := func()
            if comparefunc(result.Text, needle)
                return result
        }
        return
    }

    /**
     * Returns word clusters using a two-dimensional DBSCAN algorithm
     * @param objs An array of objects (Words, Lines etc) to cluster. Must have x, y, w, h and Text properties.
     * @param eps_x Optional epsilon value for x-axis. Default is infinite.
     * This is unused if compareFunc is provided.
     * @param eps_y Optional epsilon value for y-axis. Default is median height of objects divided by two.
     * This is unused if compareFunc is provided.
     * @param minPts Optional minimum cluster size.
     * @param compareFunc Optional comparison function to judge the minimum distance between objects
     * to consider it a cluster. Must accept two objects to compare.
     * Default comparison function determines whether the difference of middle y-coordinates of 
     * the objects are less than epsilon-y, and whether objects are less than eps_x apart on the x-axis.
     * 
     * Eg `(p1, p2) => ((Abs(p1.y+p1.h-p2.y) < 5 || Abs(p2.y+p2.h-p1.y) < 5) && ((p1.x >= p2.x && p1.x <= (p2.x+p2.w)) || ((p1.x+p1.w) >= p2.x && (p1.x+p1.w) <= (p2.x+p2.w))))`
     * will cluster objects if they are located on top of eachother on the x-axis, and less than 5 pixels
     * apart in the y-axis.
     * @param noise If provided, then will be set to an array of clusters that didn't satisfy minPts
     * @returns {Array} Array of objects with {x,y,w,h,Text,Words} properties
     */
    static Cluster(objs, eps_x:=-1, eps_y:=-1, minPts:=1, compareFunc?, &noise?) {
        local clusters := [], start := 0, cluster, word
        visited := Map(), clustered := Map(), C := [], c_n := 0, sum := 0, noise := IsSet(noise) && (noise is Array) ? noise : []
        if !IsObject(objs) || !(objs is Array)
            throw ValueError("objs argument must be an Array", -1)
        if !objs.Length
            return []
        if IsSet(compareFunc) && !HasMethod(compareFunc)
            throw ValueError("compareFunc must be a valid function", -1)

        if !IsSet(compareFunc) {
            if (eps_y < 0) {
                for point in objs
                    sum += point.h
                eps_y := (sum // objs.Length) // 2
            }
            compareFunc := (p1, p2) => Abs(p1.y+p1.h//2-p2.y-p2.h//2)<eps_y && (eps_x < 0 || (Abs(p1.x+p1.w-p2.x)<eps_x || Abs(p1.x-p2.x-p2.w)<eps_x))
        }

        ; DBSCAN adapted from https://github.com/ninopereira/DBSCAN_1D
        for point in objs {
            visited[point] := 1, neighbourPts := [], RegionQuery(point)
            if !clustered.Has(point) {
                C.Push([]), c_n += 1, C[c_n].Push(point), clustered[point] := 1
                ExpandCluster(point)
            }
            if C[c_n].Length < minPts
                noise.Push(C[c_n]), C.RemoveAt(c_n), c_n--
        }

        ; Sort clusters by x-coordinate, get cluster bounding rects, and concatenate word texts
        for cluster in C {
            OCR.SortArray(cluster,,"x")
            br := OCR.WordsBoundingRect(cluster*), br.Words := cluster, br.Text := ""
            for word in cluster
                br.Text .= word.Text " "
            br.Text := RTrim(br.Text)
            clusters.Push(br)
        }
        ; Sort clusters/lines by y-coordinate
        OCR.SortArray(clusters,,"y")
        return clusters

        ExpandCluster(P) {
            local point
            for point in neighbourPts {
                if !visited.Has(point) {
                    visited[point] := 1, RegionQuery(point)
                    if !clustered.Has(point)
                        C[c_n].Push(point), clustered[point] := 1
                }
            }
        }

        RegionQuery(P) {
            local point
            for point in objs
                if !visited.Has(point)
                    if compareFunc(P, point)
                        neighbourPts.Push(point)
        }
    }

    /**
     * Sorts an array in-place, optionally by object keys or using a callback function.
     * @param arr The array to be sorted
     * @param OptionsOrCallback Optional: either a callback function, or one of the following:
     * 
     *     N => array is considered to consist of only numeric values. This is the default option.
     *     C, C1 or COn => case-sensitive sort of strings
     *     C0 or COff => case-insensitive sort of strings
     * 
     *     The callback function should accept two parameters elem1 and elem2 and return an integer:
     *     Return integer < 0 if elem1 less than elem2
     *     Return 0 is elem1 is equal to elem2
     *     Return > 0 if elem1 greater than elem2
     * @param Key Optional: Omit it if you want to sort a array of primitive values (strings, numbers etc).
     *     If you have an array of objects, specify here the key by which contents the object will be sorted.
     * @returns {Array}
     */
    static SortArray(arr, optionsOrCallback:="N", key?) {
        static sizeofFieldType := 16 ; Same on both 32-bit and 64-bit
        if HasMethod(optionsOrCallback)
            pCallback := CallbackCreate(CustomCompare.Bind(optionsOrCallback), "F Cdecl", 2), optionsOrCallback := ""
        else {
            if InStr(optionsOrCallback, "N")
                pCallback := CallbackCreate(IsSet(key) ? NumericCompareKey.Bind(key) : NumericCompare, "F CDecl", 2)
            if RegExMatch(optionsOrCallback, "i)C(?!0)|C1|COn")
                pCallback := CallbackCreate(IsSet(key) ? StringCompareKey.Bind(key,,True) : StringCompare.Bind(,,True), "F CDecl", 2)
            if RegExMatch(optionsOrCallback, "i)C0|COff")
                pCallback := CallbackCreate(IsSet(key) ? StringCompareKey.Bind(key) : StringCompare, "F CDecl", 2)
            if InStr(optionsOrCallback, "Random")
                pCallback := CallbackCreate(RandomCompare, "F CDecl", 2)
            if !IsSet(pCallback)
                throw ValueError("No valid options provided!", -1)
        }
        mFields := NumGet(ObjPtr(arr) + (8 + (VerCompare(A_AhkVersion, "<2.1-") > 0 ? 3 : 5)*A_PtrSize), "Ptr") ; in v2.0: 0 is VTable. 2 is mBase, 3 is mFields, 4 is FlatVector, 5 is mLength and 6 is mCapacity
        DllCall("msvcrt.dll\qsort", "Ptr", mFields, "UInt", arr.Length, "UInt", sizeofFieldType, "Ptr", pCallback, "Cdecl")
        CallbackFree(pCallback)
        if RegExMatch(optionsOrCallback, "i)R(?!a)")
            this.ReverseArray(arr)
        if InStr(optionsOrCallback, "U")
            arr := this.Unique(arr)
        return arr

        CustomCompare(compareFunc, pFieldType1, pFieldType2) => (ValueFromFieldType(pFieldType1, &fieldValue1), ValueFromFieldType(pFieldType2, &fieldValue2), compareFunc(fieldValue1, fieldValue2))
        NumericCompare(pFieldType1, pFieldType2) => (ValueFromFieldType(pFieldType1, &fieldValue1), ValueFromFieldType(pFieldType2, &fieldValue2), fieldValue1 - fieldValue2)
        NumericCompareKey(key, pFieldType1, pFieldType2) => (ValueFromFieldType(pFieldType1, &fieldValue1), ValueFromFieldType(pFieldType2, &fieldValue2), fieldValue1.%key% - fieldValue2.%key%)
        StringCompare(pFieldType1, pFieldType2, casesense := False) => (ValueFromFieldType(pFieldType1, &fieldValue1), ValueFromFieldType(pFieldType2, &fieldValue2), StrCompare(fieldValue1 "", fieldValue2 "", casesense))
        StringCompareKey(key, pFieldType1, pFieldType2, casesense := False) => (ValueFromFieldType(pFieldType1, &fieldValue1), ValueFromFieldType(pFieldType2, &fieldValue2), StrCompare(fieldValue1.%key% "", fieldValue2.%key% "", casesense))
        RandomCompare(pFieldType1, pFieldType2) => (Random(0, 1) ? 1 : -1)

        ValueFromFieldType(pFieldType, &fieldValue?) {
            static SYM_STRING := 0, PURE_INTEGER := 1, PURE_FLOAT := 2, SYM_MISSING := 3, SYM_OBJECT := 5
            switch SymbolType := NumGet(pFieldType + 8, "Int") {
                case PURE_INTEGER: fieldValue := NumGet(pFieldType, "Int64") 
                case PURE_FLOAT: fieldValue := NumGet(pFieldType, "Double") 
                case SYM_STRING: fieldValue := StrGet(NumGet(pFieldType, "Ptr")+2*A_PtrSize)
                case SYM_OBJECT: fieldValue := ObjFromPtrAddRef(NumGet(pFieldType, "Ptr")) 
                case SYM_MISSING: return		
            }
        }
    }
    ; Reverses the array in-place
    static ReverseArray(arr) {
        local len := arr.Length + 1, max := (len // 2), i := 0
        while ++i <= max
            temp := arr[len - i], arr[len - i] := arr[i], arr[i] := temp
        return arr
    }
    ; Returns a new array with only unique values
    static UniqueArray(arr) {
        local unique := Map()
        for v in arr
            unique[v] := 1
        return [unique*]
    }

    ; Returns a one-dimensional array from a multi-dimensional array
    static FlattenArray(arr) {
        local r := []
        for v in arr {
            if v is Array
                r.Push(this.FlattenArray(v)*)
            else
                r.Push(v)
        }
        return r
    }

    ;; Only internal methods ahead

    static CreateDIBSection(w, h, hdc?, bpp:=32, &ppvBits:=0) {
        local hdc2 := IsSet(hdc) ? hdc : DllCall("GetDC", "Ptr", 0, "UPtr")
        , bi := Buffer(40, 0), hbm
        NumPut("int", 40, "int", w, "int", h, "ushort", 1, "ushort", bpp, "int", 0, bi)
        hbm := DllCall("CreateDIBSection", "uint", hdc2, "ptr" , bi, "uint" , 0, "uint*", &ppvBits:=0, "uint" , 0, "uint" , 0)
        if !IsSet(hdc)
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc2)
        return hbm
    }

    static CreateBitmap(X, Y, W, H, hWnd := 0, scale:=1, onlyClientArea:=0, mode:=2) {
        local sW := W*scale, sH := H*scale, flagOnlyClientArea, HDC, obm, hbm, pdc, hbm2, sW, sH
        if hWnd {
            X := 0, Y := 0, flagOnlyClientArea := onlyClientArea
            if IsObject(onlyClientArea)
                X := onlyClientArea.X, Y := onlyClientArea.Y, flagOnlyClientArea := onlyClientArea.onlyClientArea
            if mode < 2 {
                HDC := DllCall("GetDCEx", "Ptr", hWnd, "Ptr", 0, "int", 2|!flagOnlyClientArea, "Ptr")
            } else {
                hbm := this.CreateDIBSection(W, H)
                , hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "UPtr")
                , obm := DllCall("SelectObject", "Ptr", HDC, "Ptr", HBM)
                , DllCall("PrintWindow", "ptr", hwnd, "ptr", hdc, "uint", 2|!!flagOnlyClientArea)
                if scale != 1 {
                    PDC := DllCall("CreateCompatibleDC", "Ptr", HDC, "UPtr")
                    , hbm2 := DllCall("CreateCompatibleBitmap", "Ptr", HDC, "Int", sW, "Int", sH, "UPtr")
                    , obm2 := DllCall("SelectObject", "Ptr", PDC, "Ptr", HBM2)
                    , DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", HDC, "Int", X, "Int", Y, "Int", W, "Int", H, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY
                    , DllCall("SelectObject", "Ptr", PDC, "Ptr", obm2)
                    , DllCall("DeleteDC", "Ptr", PDC)
                    , DllCall("DeleteObject", "UPtr", HBM)
                    , hbm := hbm2
                }
                DllCall("SelectObject", "Ptr", HDC, "Ptr", obm)
                DllCall("DeleteDC", "Ptr", HDC)
                return this.IBase(HBM).DefineProp("__Delete", {call:(*)=>DllCall("DeleteObject", "UPtr", HBM)})
            }
        } else {
            HDC := DllCall("GetDC", "Ptr", 0, "UPtr")
        }
        HBM := DllCall("CreateCompatibleBitmap", "Ptr", HDC, "Int", Max(40,sW), "Int", Max(40,sH), "UPtr")
        , PDC := DllCall("CreateCompatibleDC", "Ptr", HDC, "UPtr")
        , OBM := DllCall("SelectObject", "Ptr", PDC, "Ptr", HBM)
        if sW < 40 || sH < 40 ; Fills the bitmap so it's at least 40x40, which seems to improve recognition
            DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", Max(40,sW), "Int", Max(40,sH), "Ptr", HDC, "Int", X, "Int", Y, "Int", 1, "Int", 1, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY. 
        DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", HDC, "Int", X, "Int", Y, "Int", W, "Int", H, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY
        , DllCall("SelectObject", "Ptr", PDC, "Ptr", OBM)
        , DllCall("DeleteDC", "Ptr", PDC)
        , DllCall("ReleaseDC", "Ptr", 0, "Ptr", HDC)
        return this.IBase(HBM).DefineProp("__Delete", {call:(*)=>DllCall("DeleteObject", "UPtr", HBM)})
    }

    static HBitmapToRandomAccessStream(hBitmap) {
        static PICTYPE_BITMAP := 1
             , BSOS_DEFAULT   := 0
             , sz := 8 + A_PtrSize*2
        local PICTDESC, riid, size, pIRandomAccessStream
             
        DllCall("Ole32\CreateStreamOnHGlobal", "Ptr", 0, "UInt", true, "Ptr*", pIStream:=this.IBase(), "UInt")
        , PICTDESC := Buffer(sz, 0)
        , NumPut("uint", sz, "uint", PICTYPE_BITMAP, "ptr", IsInteger(hBitmap) ? hBitmap : hBitmap.ptr, PICTDESC)
        , riid := this.CLSIDFromString(this.IID_IPicture)
        , DllCall("OleAut32\OleCreatePictureIndirect", "Ptr", PICTDESC, "Ptr", riid, "UInt", 0, "Ptr*", pIPicture:=this.IBase(), "UInt")
        , ComCall(15, pIPicture, "Ptr", pIStream, "UInt", true, "uint*", &size:=0, "UInt") ; IPicture::SaveAsFile
        , riid := this.CLSIDFromString(this.IID_IRandomAccessStream)
        , DllCall("ShCore\CreateRandomAccessStreamOverStream", "Ptr", pIStream, "UInt", BSOS_DEFAULT, "Ptr", riid, "Ptr*", pIRandomAccessStream:=this.IBase(), "UInt")
        Return pIRandomAccessStream
    }

    static DisplayHBitmap(hBitmap, W:=640, H:=640) {
        local gImage := Gui()
        , hPic := gImage.Add("Text", "0xE w" W " h" H)
        SendMessage(0x172, 0, hBitmap,, hPic.Hwnd)
        gImage.Show()
        WinWaitClose gImage
    }

    static CreateClass(str, interface?) {
        local hString := this.CreateHString(str), result
        if !IsSet(interface) {
            result := DllCall("Combase.dll\RoActivateInstance", "ptr", hString, "ptr*", cls:=this.IBase(), "uint")
        } else {
            GUID := this.CLSIDFromString(interface)
            result := DllCall("Combase.dll\RoGetActivationFactory", "ptr", hString, "ptr", GUID, "ptr*", cls:=this.IBase(), "uint")
        }
        if (result != 0) {
            if (result = 0x80004002)
                throw Error("No such interface supported", -1, interface)
            else if (result = 0x80040154)
                throw Error("Class not registered", -1)
            else
                throw Error(result)
        }
        this.DeleteHString(hString)
        return cls
    }
    
    static CreateHString(str) => (DllCall("Combase.dll\WindowsCreateString", "wstr", str, "uint", StrLen(str), "ptr*", &hString:=0), hString)
    
    static DeleteHString(hString) => DllCall("Combase.dll\WindowsDeleteString", "ptr", hString)
    
    static WaitForAsync(&obj) {
        local AsyncInfo := ComObjQuery(obj, this.IID_IAsyncInfo), status, ErrorCode
        Loop {
            ComCall(7, AsyncInfo, "uint*", &status:=0)   ; IAsyncInfo.Status
            if (status != 0) {
                if (status != 1) {
                    ComCall(8, ASyncInfo, "uint*", &ErrorCode:=0)   ; IAsyncInfo.ErrorCode
                    throw Error("AsyncInfo failed with status error " ErrorCode, -1)
                }
             break
          }
          Sleep 10
        }
        ComCall(8, obj, "ptr*", ObjectResult:=this.IBase())   ; GetResults
        obj := ObjectResult
    }

    static CloseIClosable(pClosable) {
        static IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}"
        local Close := ComObjQuery(pClosable, IClosable)
        ComCall(6, Close)   ; Close
        if !IsObject(pClosable)
            ObjRelease(pClosable)
    }

    static CLSIDFromString(IID) {
        local CLSID := Buffer(16), res
        if res := DllCall("ole32\CLSIDFromString", "WStr", IID, "Ptr", CLSID, "UInt")
           throw Error("CLSIDFromString failed. Error: " . Format("{:#x}", res))
        Return CLSID
    }

    static NormalizeCoordinates(result, scale) {
        local word
        if scale != 1 {
            for word in result.Words
                word.x := Integer(word.x / scale), word.y := Integer(word.y / scale), word.w := Integer(word.w / scale), word.h := Integer(word.h / scale), word.BoundingRect := {X:word.x, Y:word.y, W:word.w, H:word.h}
        }
        return result
    }
}
