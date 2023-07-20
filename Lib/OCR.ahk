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
 * OCR.FromBitmap(HBitmap, lang?)
 * 
 * Additional methods:
 * OCR.GetAvailableLanguages()
 * OCR.LoadLanguage(lang:="FirstFromAvailableLanguages")
 * OCR.WaitText(needle, timeout:=-1, func?, casesense:=False, comparefunc?)
 *      Calls a func (the provided OCR method) until a string is found
 * OCR.WordsBoundingRect(words*)
 *      Returns the bounding rectangle for multiple words
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
 * Result.Highlight(obj?, showTime:=2000, color:="Red", d:=2)
 *      Highlights a Word, Line, or object with {x,y,w,h} properties on the screen, or removes the highlighting
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
        this.LanguageFactory := OCR.CreateClass("Windows.Globalization.Language", ILanguageFactory := "{9B0252AC-0C27-44F8-B792-9793FB66C63E}")
        this.BitmapTransform := OCR.CreateClass("Windows.Graphics.Imaging.BitmapTransform")
        this.BitmapDecoderStatics := OCR.CreateClass("Windows.Graphics.Imaging.BitmapDecoder", IBitmapDecoderStatics := "{438CCB26-BCEF-4E95-BAD6-23A822E58D01}")
        this.OcrEngineStatics := OCR.CreateClass("Windows.Media.Ocr.OcrEngine", IOcrEngineStatics := "{5BFFA85A-3384-3540-9940-699120D428A8}")
        ComCall(6, this.OcrEngineStatics, "uint*", &MaxImageDimension:=0)   ; MaxImageDimension
        this.MaxImageDimension := MaxImageDimension
    }

    /**
     * Returns an OCR results object for an IRandomAccessStream.
     * Images of other types should be first converted to this format (eg from file, from bitmap).
     * @param pIRandomAccessStream Pointer or an object containing a ptr to the stream
     * @param {String} lang OCR language. Default is first from available languages.
     * @returns {Ocr} 
     */
    __New(pIRandomAccessStream?, lang := "FirstFromAvailableLanguages") {
        if IsSet(lang) || !OCR.HasOwnProp("CurrentLanguage")
            OCR.LoadLanguage(lang?)
        ComCall(14, OCR.BitmapDecoderStatics, "ptr", pIRandomAccessStream, "ptr*", BitmapDecoder:=OCR.IBase())   ; CreateAsync
        OCR.WaitForAsync(&BitmapDecoder)
        BitmapFrame := ComObjQuery(BitmapDecoder, IBitmapFrame := "{72A49A1C-8081-438D-91BC-94ECFC8185C6}")
        ComCall(12, BitmapFrame, "uint*", &width:=0)   ; get_PixelWidth
        ComCall(13, BitmapFrame, "uint*", &height:=0)   ; get_PixelHeight
        if (width > OCR.MaxImageDimension) or (height > OCR.MaxImageDimension)
           throw ValueError("Image is too big - " width "x" height ".`nIt should be maximum - " OCR.MaxImageDimension " pixels")

        BitmapFrameWithSoftwareBitmap := ComObjQuery(BitmapDecoder, IBitmapFrameWithSoftwareBitmap := "{FE287C9A-420C-4963-87AD-691436E08383}")
        if width < 40 || height < 40 {
            scale := 40.0 / Min(width, height), this.ImageWidth := Ceil(width*scale), this.ImageHeight := Ceil(height*scale)
            ComCall(7, this.BitmapTransform, "int", this.ImageWidth) ; put_ScaledWidth
            ComCall(9, this.BitmapTransform, "int", this.ImageHeight) ; put_ScaledHeight
            ComCall(8, BitmapFrame, "uint*", &BitmapPixelFormat:=0) ; get_BitmapPixelFormat
            ComCall(9, BitmapFrame, "uint*", &BitmapAlphaMode:=0) ; get_BitmapAlphaMode
            ComCall(8, BitmapFrameWithSoftwareBitmap, "uint", BitmapPixelFormat, "uint", BitmapAlphaMode, "ptr", this.BitmapTransform, "uint", IgnoreExifOrientation := 0, "uint", DoNotColorManage := 0, "ptr*", SoftwareBitmap:=OCR.IBase()) ; GetSoftwareBitmapAsync
        } else {
            this.ImageWidth := width, this.ImageHeight := height
            ComCall(6, BitmapFrameWithSoftwareBitmap, "ptr*", SoftwareBitmap:=OCR.IBase())   ; GetSoftwareBitmapAsync
        }
        OCR.WaitForAsync(&SoftwareBitmap)

        ComCall(6, OCR.OcrEngine, "ptr", SoftwareBitmap, "ptr*", OcrResult:=OCR.IBase())   ; RecognizeAsync
        OCR.WaitForAsync(&OcrResult)

        ; Cleanup
        OCR.CloseIClosable(pIRandomAccessStream)
        OCR.CloseIClosable(SoftwareBitmap)

        this.ptr := OcrResult.ptr, ObjAddRef(OcrResult.ptr)
    }
    __Delete() => this.ptr ? ObjRelease(this.ptr) : 0

    ; Gets the recognized text.
    Text {
        get {
            ComCall(8, this, "ptr*", &hAllText:=0)   ; get_Text
            buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hAllText, "uint*", &length:=0, "ptr")
            this.DefineProp("Text", {Value:StrGet(buf, "UTF-16")})
            OCR.DeleteHString(hAllText)
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
            ComCall(6, this, "ptr*", LinesList:=OCR.IBase()) ; get_Lines
            ComCall(7, LinesList, "int*", &count:=0) ; count
            lines := []
            loop count {
                ComCall(6, LinesList, "int", A_Index-1, "ptr*", OcrLine:=OCR.OCRLine())               
                lines.Push(OcrLine)
            }
            this.DefineProp("Lines", {Value:lines})
            return lines
        }
    }

    ; Returns all Word objects for the result. Equivalent to looping over all the Lines and getting the Words.
    Words {
        get {
            words := []
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
     * otherwise the Relative properties values will be added to the x and y coordinates as offsets.
     */
    Click(Obj, WhichButton?, ClickCount?, DownOrUp?) {
        if !obj.HasOwnProp("x") && InStr(Type(obj), "OCR")
            obj := OCR.WordsBoundingRect(obj.Words)
        x := obj.x, y := obj.y, w := obj.w, h := obj.h
        if this.HasOwnProp("Relative") {
            if this.Relative.HasOwnProp("Hwnd") {
                if !WinActive(this.Relative.Hwnd) {
                    WinActivate(this.Relative.Hwnd)
                    WinWaitActive(this.Relative.Hwnd,,1)
                }
            } else
                x += this.Relative.x, y += this.Relative.y
        }
        oldCoordMode := A_CoordModeMouse
        CoordMode "Mouse", this.HasOwnProp("Relative") && this.Relative.HasOwnProp("Type") ? this.Relative.Type : "Screen"
        Click(x+w//2, y+h//2, WhichButton?, ClickCount?, DownOrUp?)
        CoordMode "Mouse", oldCoordMode
    }

    /**
     * ControlClicks an object
     * @param obj The object to click, which can be a OCR result object, Line, Word, or Object {x,y,w,h}
     * If this object (the one Click is called from) contains a "Relative" property (this is
     * added by default with OCR.FromWindow) containing a Hwnd property, then that window will be activated,
     * otherwise the Relative properties values will be added to the x and y coordinates as offsets.
     * @param WinTitle If WinTitle is set, then the coordinates stored in Obj will be converted to
     * client coordinates and ControlClicked.
     */
    ControlClick(obj, WinTitle?, WinText?, WhichButton?, ClickCount?, Options?, ExcludeTitle?, ExcludeText?) {
        if !obj.HasOwnProp("x") && InStr(Type(obj), "OCR")
            obj := OCR.WordsBoundingRect(obj.Words)
        x := obj.x, y := obj.y, w := obj.w, h := obj.h
        if this.HasOwnProp("Relative") && this.Relative.HasOwnProp("Type") {
            hWnd := this.Relative.hWnd
            if this.Relative.Type = "Window" {
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
     * * Unset - if highlighting exists then removes the highlighting
     * * 0 - Indefinite highlighting
     * * Positive integer (eg 2000) - will highlight and pause for the specified amount of time in ms
     * * Negative integer - will highlight for the specified amount of time in ms, but script execution will continue
     * @param {string} color The color of the highlighting. Default is red.
     * @param {number} d The border thickness of the highlighting in pixels. Default is 2.
     * @returns {OCR}
     */
    Highlight(obj?, showTime:=2000, color:="Red", d:=2) {
        static guis := []
        if !IsSet(obj) {
            for _, r in guis
                r.Destroy()
            guis := []
            return this
        }
        if !guis.Length {
            Loop 4
                guis.Push(Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000000"))
        }
        if Type(obj) = "OCR.OCRLine" || Type(obj) = "OCR"
            obj := OCR.WordsBoundingRect(obj.Words*)
        x := obj.x, y := obj.y, w := obj.w, h := obj.h
        if this.HasOwnProp("Relative")
            x += this.Relative.x, y += this.Relative.y

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
            this.Highlight()
        } else if showTime < 0
            SetTimer(this.GetMethod("Highlight"), -Abs(showTime))
        return this
    }

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
                                result := OCR.WordsBoundingRect(found*)
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
                            result := OCR.WordsBoundingRect(found*)
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
     * Crops the result object to contain only results from an area defined by points (x1,y1) and (x2,y2).
     * Note that these coordinates are relative to the result object, not to the screen.
     * @param {Integer} x1 x coordinate of the top left corner of the search area
     * @param {Integer} y1 y coordinate of the top left corner of the search area
     * @param {Integer} x2 x coordinate of the bottom right corner of the search area
     * @param {Integer} y2 y coordinate of the bottom right corner of the search area
     * @returns {OCR}
     */
    Crop(x1:=-100000, y1:=-100000, x2:=100000, y2:=100000) {
        result := this.Clone()
        croppedLines := [], croppedText := ""
        for line in result.Lines {
            croppedWords := [], lineText := ""
            for word in line.Words {
                if word.x >= x1 && word.y >= y1 && (word.x+word.w) <= x2 && (word.y+word.h) <= y2
                    croppedWords.Push(word), lineText .= word.Text " ", ObjAddRef(word.ptr)
            }
            if croppedWords.Length {
                line := {Text:Trim(lineText), Words:croppedWords}
                line.base.__Class := "OCR.OCRLine"
                croppedLines.Push(line)
                croppedText .= lineText
            }
        }
        result.DefineProp("Lines", {Value:croppedLines})
        result.DefineProp("Text", {Value:Trim(croppedText)})
        result.DefineProp("Words", OCR.Prototype.GetOwnPropDesc("Words"))
        return result
    }

    class OCRLine extends OCR.IBase {
        ; Gets the recognized text for the line.
        Text {
            get {
                ComCall(7, this, "ptr*", &hText:=0)   ; get_Text
                buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length:=0, "ptr")
                text := StrGet(buf, "UTF-16")
                OCR.DeleteHString(hText)
                this.DefineProp("Text", {Value:text})
                return text
            }
        }

        ; Gets the Word objects for the line
        Words {
            get {
                ComCall(6, this, "ptr*", WordsList:=OCR.IBase())   ; get_Words
                ComCall(7, WordsList, "int*", &WordsCount:=0)   ; Words count
                words := []
                loop WordsCount {
                   ComCall(6, WordsList, "int", A_Index-1, "ptr*", OcrWord:=OCR.OCRWord())
                   words.Push(OcrWord)
                }
                this.DefineProp("Words", {Value:words})
                return words
            }
        }
    }

    class OCRWord extends OCR.IBase {
        ; Gets the recognized text for the word
        Text {
            get {
                ComCall(7, this, "ptr*", &hText:=0)   ; get_Text
                buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length:=0, "ptr")
                text := StrGet(buf, "UTF-16")
                OCR.DeleteHString(hText)
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
         GUID := OCR.CLSIDFromString(OCR.IID_IRandomAccessStream)
         DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", FileName, "uint", Read := 0, "ptr", GUID, "ptr*", IRandomAccessStream:=OCR.IBase())
         return OCR(IRandomAccessStream, lang?)
    }

    /**
     * Returns an OCR results object for a given window. Locations of the words will be relative to the
     * window or client area, so for interactions use CoordMode "Window" or "Client".
     * @param WinTitle A window title or other criteria identifying the target window.
     * @param lang OCR language. Default is first from available languages.
     * @param scale The scaling factor to use.
     * @param {Number} onlyClientArea Whether only the client area or the whole window should be OCR-d
     * @param {Number} mode Different methods of capturing the window. 0 = uses GetDC with BitBlt, 2 = uses PrintWindow. 
     * Add 1 to make a transparent window totally opaque. 
     * @returns {Ocr} 
     */
    static FromWindow(WinTitle:="", lang?, scale:=1, onlyClientArea:=0, mode:=2) {
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
        If onlyClientArea {
            DllCall("GetClientRect", "ptr", hwnd, "ptr", rc:=Buffer(16))
            W := NumGet(rc, 8, "int"), H := NumGet(rc, 12, "int")
            pt:=Buffer(8, 0), NumPut("int64", 0, pt)
            DllCall("ClientToScreen", "Ptr", hwnd, "Ptr", pt)
            X:=NumGet(pt,"int"), Y:=NumGet(pt,4,"int")
        } else {
            rect := Buffer(16, 0)
            DllCall("GetWindowRect", "UPtr", hwnd, "Ptr", rect, "UInt")
            X := NumGet(rect, 0, "Int"), Y := NumGet(rect, 4, "Int")
            x2 := NumGet(rect, 8, "Int"), y2 := NumGet(rect, 12, "Int")
            W := Abs(Max(X, X2) - Min(X, X2))
            H := Abs(Max(Y, Y2) - Min(Y, Y2))
        }
        hBitMap := OCR.CreateBitmap(X, Y, W, H, hWnd, scale, onlyClientArea, mode)
        ;OCR.DisplayHBitmap(hBitMap)
        if mode&1
            WinSetExStyle(oldStyle, hwnd)
        result := OCR(OCR.HBitmapToRandomAccessStream(hBitMap), lang?)
        result.Relative := {X:X, Y:Y, Type:(onlyClientArea ? "Client" : "Window"), Hwnd:hWnd}
        OCR.NormalizeCoordinates(result, scale)
        return result
    }

    /**
     * Returns an OCR results object for the whole desktop. Locations of the words will be relative to
     * the screen (CoordMode "Screen")
     * @param lang OCR language. Default is first from available languages.
     * @returns {Ocr} 
     */
    static FromDesktop(lang?, scale:=1) => OCR.FromRect(0, 0, A_ScreenWidth, A_ScreenHeight, lang?, scale)

    /**
     * Returns an OCR results object for a region of the screen. Locations of the words will be relative
     * to the top left corner of the rectangle.
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
        hBitmap := OCR.CreateBitmap(X, Y, W, H,,scale)
        result := OCR(OCR.HBitmapToRandomAccessStream(hBitmap), lang?)
        result.Relative := {x:x, y:y}
        return OCR.NormalizeCoordinates(result, scale)
    }

    /**
     * Returns an OCR results object from a hBitmap object. Locations of the words will be relative
     * to the top left corner of the bitmap.
     * @param hBitmap An hBitmap pointer or an object with a ptr property
     * @param lang OCR language. Default is first from available languages.
     * @returns {ocr} 
     */
    static FromBitmap(hBitmap, lang?) => OCR(OCR.HBitmapToRandomAccessStream(hBitmap), lang?)

    /**
     * Returns all available languages as a string, where the languages are separated by newlines.
     * @returns {String} 
     */
    static GetAvailableLanguages() {
        static GlobalizationPreferencesStatics
        if !IsSet(GlobalizationPreferencesStatics)
            GlobalizationPreferencesStatics := OCR.CreateClass("Windows.System.UserProfile.GlobalizationPreferences", IGlobalizationPreferencesStatics := "{01BF4326-ED37-4E96-B0E9-C1340D1EA158}")
        ComCall(9, GlobalizationPreferencesStatics, "ptr*", &LanguageList:=0)   ; get_Languages
        ComCall(7, LanguageList, "int*", &count:=0)   ; count
        Loop count {
            ComCall(6, LanguageList, "int", A_Index-1, "ptr*", &hString:=0)   ; get_Item
            ComCall(6, this.LanguageFactory, "ptr", hString, "ptr*", &LanguageTest:=0)   ; CreateLanguage
            ComCall(8, this.OcrEngineStatics, "ptr", LanguageTest, "int*", &bool:=0)   ; IsLanguageSupported
            if (bool = 1) {
                ComCall(6, LanguageTest, "ptr*", &hText:=0)
                buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length:=0, "ptr")
                text .= StrGet(buf, "UTF-16") "`n"
            }
            ObjRelease(LanguageTest)
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
        if this.HasOwnProp("CurrentLanguage") && this.HasOwnProp("OcrEngine") && this.CurrentLanguage = lang
            return
        if (lang = "FirstFromAvailableLanguages")
            ComCall(10, this.OcrEngineStatics, "ptr*", OcrEngine:=OCR.IBase())   ; TryCreateFromUserProfileLanguages
        else {
            hString := OCR.CreateHString(lang)
            ComCall(6, this.LanguageFactory, "ptr", hString, "ptr*", Language:=OCR.IBase())   ; CreateLanguage
            OCR.DeleteHString(hString)
            ComCall(9, this.OcrEngineStatics, "ptr", Language, "ptr*", OcrEngine:=OCR.IBase())   ; TryCreateFromLanguage
        }
        if (OcrEngine.ptr = 0)
            Throw Error("Can not use language `"" lang "`" for OCR, please install language pack.")
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
        X1 := 100000000, Y1 := 100000000, X2 := -100000000, Y2 := -100000000
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
        endTime := A_TickCount+timeout
        if !IsSet(func)
            func := OCR.FromDesktop
        if !IsSet(comparefunc)
            comparefunc := InStr.Bind(,,casesense)
        While timeout > 0 ? (A_TickCount < endTime) : 1 {
            result := func()
            if comparefunc(result.Text, needle)
                return result
        }
        return
    }

    ;; Only internal methods ahead

    static CreateDIBSection(w, h, hdc?, bpp:=32, &ppvBits:=0) {
        hdc2 := IsSet(hdc) ? hdc : DllCall("GetDC", "Ptr", 0, "UPtr")
        bi := Buffer(40, 0)
        NumPut("int", 40, "int", w, "int", h, "ushort", 1, "ushort", bpp, "int", 0, bi)
        hbm := DllCall("CreateDIBSection", "uint", hdc2, "ptr" , bi, "uint" , 0, "uint*", &ppvBits:=0, "uint" , 0, "uint" , 0)
        if !IsSet(hdc)
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc2)
        return hbm
    }

    static CreateBitmap(X, Y, W, H, hWnd := 0, scale:=1, onlyClientArea:=0, mode:=2) {
        static CAPTUREBLT
        if !IsSet(CAPTUREBLT) {
            DllCall("Dwmapi\DwmIsCompositionEnabled", "Int*", &compositionEnabled:=0)
            CAPTUREBLT:= compositionEnabled ? 0 : 0x40000000
        }
        sW := W*scale, sH := H*scale
        if hWnd {
            if mode < 2 {
                X := 0, Y := 0
                HDC := DllCall("GetDCEx", "Ptr", hWnd, "Ptr", 0, "int", 2|!onlyClientArea, "Ptr")
            } else {
                hbm := OCR.CreateDIBSection(W, H)
                hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "UPtr")
                obm := DllCall("SelectObject", "Ptr", HDC, "Ptr", HBM)
                DllCall("PrintWindow", "uint", hwnd, "uint", hdc, "uint", 2|!!onlyClientArea)
                if scale != 1 {
                    PDC := DllCall("CreateCompatibleDC", "Ptr", HDC, "UPtr")
                    hbm2 := DllCall("CreateCompatibleBitmap", "Ptr", HDC, "Int", sW, "Int", sH, "UPtr")
                    DllCall("SelectObject", "Ptr", PDC, "Ptr", HBM2)
                    DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", HDC, "Int", 0, "Int", 0, "Int", W, "Int", H, "UInt", 0x00CC0020 | CAPTUREBLT) ; SRCCOPY
                    DllCall("DeleteDC", "Ptr", PDC)
                    DllCall("DeleteObject", "UPtr", HBM)
                    hbm := hbm2
                }
                DllCall("DeleteDC", "Ptr", HDC)
                return OCR.IBase(HBM).DefineProp("__Delete", {call:(*)=>DllCall("DeleteObject", "UPtr", HBM)})
            }
        } else {
            HDC := DllCall("GetDC", "Ptr", 0, "UPtr")
        }
        HBM := DllCall("CreateCompatibleBitmap", "Ptr", HDC, "Int", Max(40,sW), "Int", Max(40,sH), "UPtr")
        PDC := DllCall("CreateCompatibleDC", "Ptr", HDC, "UPtr")
        DllCall("SelectObject", "Ptr", PDC, "Ptr", HBM)
        if sW < 40 || sH < 40 ; Fills the bitmap so it's at least 40x40, which seems to improve recognition
            DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", Max(40,sW), "Int", Max(40,sH), "Ptr", HDC, "Int", X, "Int", Y, "Int", 1, "Int", 1, "UInt", 0x00CC0020 | CAPTUREBLT) ; SRCCOPY. 
        DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", HDC, "Int", X, "Int", Y, "Int", W, "Int", H, "UInt", 0x00CC0020 | CAPTUREBLT) ; SRCCOPY
        DllCall("DeleteDC", "Ptr", PDC)
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", HDC)
        return OCR.IBase(HBM).DefineProp("__Delete", {call:(*)=>DllCall("DeleteObject", "UPtr", HBM)})
    }

    static HBitmapToRandomAccessStream(hBitmap) {
        static PICTYPE_BITMAP := 1
             , BSOS_DEFAULT   := 0
             , sz := 8 + A_PtrSize*2
             
        DllCall("Ole32\CreateStreamOnHGlobal", "Ptr", 0, "UInt", true, "Ptr*", pIStream:=OCR.IBase(), "UInt")
        
        PICTDESC := Buffer(sz, 0)
        NumPut("uint", sz, "uint", PICTYPE_BITMAP, "ptr", IsInteger(hBitmap) ? hBitmap : hBitmap.ptr, PICTDESC)
        riid := OCR.CLSIDFromString(OCR.IID_IPicture)
        DllCall("OleAut32\OleCreatePictureIndirect", "Ptr", PICTDESC, "Ptr", riid, "UInt", 0, "Ptr*", pIPicture:=OCR.IBase(), "UInt")
        ; IPicture::SaveAsFile
        ComCall(15, pIPicture, "Ptr", pIStream, "UInt", true, "uint*", &size:=0, "UInt")
        riid := OCR.CLSIDFromString(OCR.IID_IRandomAccessStream)
        DllCall("ShCore\CreateRandomAccessStreamOverStream", "Ptr", pIStream, "UInt", BSOS_DEFAULT, "Ptr", riid, "Ptr*", pIRandomAccessStream:=OCR.IBase(), "UInt")
        Return pIRandomAccessStream
    }

    static DisplayHBitmap(hBitmap, W:=640, H:=640) {
        gImage := Gui()
        hPic := gImage.Add("Text", "0xE w" W " h" H)
        SendMessage(0x172, 0, hBitmap,, hPic.Hwnd)
        gImage.Show()
        WinWaitClose gImage
    }

    static CreateClass(str, interface?) {
        hString := OCR.CreateHString(str)
        if !IsSet(interface) {
            result := DllCall("Combase.dll\RoActivateInstance", "ptr", hString, "ptr*", cls:=OCR.IBase(), "uint")
        } else {
            GUID := OCR.CLSIDFromString(interface)
            result := DllCall("Combase.dll\RoGetActivationFactory", "ptr", hString, "ptr", GUID, "ptr*", cls:=OCR.IBase(), "uint")
        }
        if (result != 0) {
            if (result = 0x80004002)
                throw Error("No such interface supported", -1, interface)
            else if (result = 0x80040154)
                throw Error("Class not registered", -1)
            else
                throw Error(result)
        }
        OCR.DeleteHString(hString)
        return cls
    }
    
    static CreateHString(str) => (DllCall("Combase.dll\WindowsCreateString", "wstr", str, "uint", StrLen(str), "ptr*", &hString:=0), hString)
    
    static DeleteHString(hString) => DllCall("Combase.dll\WindowsDeleteString", "ptr", hString)
    
    static WaitForAsync(&obj) {
        AsyncInfo := ComObjQuery(obj, OCR.IID_IAsyncInfo)
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
        ComCall(8, obj, "ptr*", ObjectResult:=OCR.IBase())   ; GetResults
        obj := ObjectResult
    }

    static CloseIClosable(pClosable) {
        static IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}"
        Close := ComObjQuery(pClosable, IClosable)
        ComCall(6, Close)   ; Close
        if !IsObject(pClosable)
            ObjRelease(pClosable)
    }

    static CLSIDFromString(IID) {
        CLSID := Buffer(16)
        if res := DllCall("ole32\CLSIDFromString", "WStr", IID, "Ptr", CLSID, "UInt")
           throw Error("CLSIDFromString failed. Error: " . Format("{:#x}", res))
        Return CLSID
    }

    static NormalizeCoordinates(result, scale) {
        if scale != 1 {
            for word in result.Words
                word.x := Integer(word.x / scale), word.y := Integer(word.y / scale), word.w := Integer(word.w / scale), word.h := Integer(word.h / scale), word.BoundingRect := {X:word.x, Y:word.y, W:word.w, H:word.h}
        }
        return result
    }
}