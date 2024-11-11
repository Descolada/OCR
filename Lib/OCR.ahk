#Requires AutoHotkey v2

/**
 * OCR library: a wrapper for the the UWP Windows.Media.Ocr library.
 * Based on the UWP OCR function for AHK v1 by malcev.
 * 
 * Ways of initiating OCR:
 * OCR(RandomAccessStreamOrSoftwareBitmap, lang:="FirstFromAvailableLanguages", transform:=1, decoder?)
 * OCR.FromDesktop(lang?, transform:=1)
 * OCR.FromRect(X, Y, W, H, lang?, transform:=1)
 * OCR.FromWindow(WinTitle?, lang?, transform:=1, onlyClientArea:=0, mode:=4)
 * OCR.FromFile(FileName, lang?, transform:=1)
 * OCR.FromBitmap(bitmap, lang?, transform:=1, hDC?)
 * OCR.FromPDF(FileName, lang?, transform:=1, start:=1, end?)   => returns an array of results for each PDF page
 * OCR.FromPDFPage(FileName, page:=1, lang?, transform:=1)
 * 
 * Note: the first parameter of the OCR initiation methods may be an object mimicking named parameters.
 * Eg. OCR.FromDesktop({lang:"en-us", grayscale:1})
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
 * OCR.PerformanceMode
 *      Increases speed of OCR acquisition by about 20-50ms if set to 1, but also increases CPU usage. Default is 0. 
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
 * Line.x,y,w,h      => Size and location of the Line. Coordinates are relative to the original image.
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
         , IID_IAsyncOperation_OcrResult        := "{c7d7118e-ae36-59c0-ac76-7badee711c8b}"
         , IID_IAsyncOperation_SoftwareBitmap   := "{c4a10980-714b-5501-8da2-dbdacce70f73}"
         , IID_IAsyncOperation_BitmapDecoder    := "{aa94d8e9-caef-53f6-823d-91b6e8340510}"
         , IID_IAsyncOperationCompletedHandler_OcrResult        := "{989c1371-444a-5e7e-b197-9eaaf9d2829a}"
         , IID_IAsyncOperationCompletedHandler_SoftwareBitmap   := "{b699b653-33ed-5e2d-a75f-02bf90e32619}"
         , IID_IAsyncOperationCompletedHandler_BitmapDecoder    := "{bb6514f2-3cfb-566f-82bc-60aabd302d53}"
         , IID_IPdfDocumentStatics := "{433A0B5F-C007-4788-90F2-08143D922599}"
         , Vtbl_GetDecoder := {bmp:6, jpg:7, jpeg:7, png:8, tiff:9, gif:10, jpegxr:11, ico:12}
         , PerformanceMode := 0
         , DisplayImage := 0

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
        this.SoftwareBitmapFactory := this.CreateClass("Windows.Graphics.Imaging.SoftwareBitmap", "{c99feb69-2d62-4d47-a6b3-4fdb6a07fdf8}")
        this.BitmapTransform := this.CreateClass("Windows.Graphics.Imaging.BitmapTransform")
        this.BitmapDecoderStatics := this.CreateClass("Windows.Graphics.Imaging.BitmapDecoder", IBitmapDecoderStatics := "{438CCB26-BCEF-4E95-BAD6-23A822E58D01}")
        this.BitmapEncoderStatics := this.CreateClass("Windows.Graphics.Imaging.BitmapEncoder", IBitmapDecoderStatics := "{a74356a7-a4e4-4eb9-8e40-564de7e1ccb2}")
        this.SoftwareBitmapStatics := this.CreateClass("Windows.Graphics.Imaging.SoftwareBitmap", ISoftwareBitmapStatics := "{df0385db-672f-4a9d-806e-c2442f343e86}")
        this.OcrEngineStatics := this.CreateClass("Windows.Media.Ocr.OcrEngine", IOcrEngineStatics := "{5BFFA85A-3384-3540-9940-699120D428A8}")
        ComCall(6, this.OcrEngineStatics, "uint*", &MaxImageDimension:=0)   ; MaxImageDimension
        this.MaxImageDimension := MaxImageDimension
        DllCall("Dwmapi\DwmIsCompositionEnabled", "Int*", &compositionEnabled:=0)
        this.CAPTUREBLT := compositionEnabled ? 0 : 0x40000000
        /*  // Based on code by AHK forums user Xtra
            unsigned int Convert_GrayScale(unsigned int bitmap[], unsigned int w, unsigned int h, unsigned int Stride)
            {
                unsigned int a, r, g, b, gray, ARGB;
                unsigned int x, y, offset = Stride/4;
                for (y = 0; y < h; ++y) {
                    for (x = 0; x < w; ++x) {
                        ARGB = bitmap[x+(y*offset)];
                        a = ARGB & 0xFF000000;
                        r = (ARGB & 0x00FF0000) >> 16;
                        g = (ARGB & 0x0000FF00) >> 8;
                        b = (ARGB & 0x000000FF);
                        gray = ((300 * r) + (590 * g) + (110 * b)) >> 10;
                        bitmap[x+(y*offset)] = (gray << 16) | (gray << 8) | gray | a;
                    }
                }
                return 0;
            }
         */
        this.GrayScaleMCode := this.MCode((A_PtrSize = 4) 
        ? "2,x86:VVdWU4PsCIt0JCiLVCQki0QkIMHuAok0JIXSD4SDAAAAhcB0f408tQAAAAAx9ol8JASLfCQcjRyHMf+NdCYAkItEJByNDLiNtCYAAAAAZpCLEYPBBInQD7buwegQae1OAgAAD7bAacAsAQAAAegPtuqB4gAAAP9r7W4B6MHoConFCcLB4AjB5RAJ6gnQiUH8Odl1vIPGAQM8JANcJAQ5dCQkdZyDxAgxwFteX13D" 
        : "2,x64:QVZVV1ZTRInOSYnLQYnSRYnGwe4CRYXAdHJFMclFMcCF0nRoDx9AAESJyg8fRAAAidCDwgFJjQyDizmJ+In7wegQD7bvD7bAae1OAgAAacAsAQAAAehAD7bvgecAAAD/a+1uAejB6AqJxQnHweAIweUQCe8Jx4k5RDnSdbNBg8ABQQHxQQHyRTnGdZwxwFteX11BXsM=")
        /*
            unsigned int Invert_Colors(unsigned int bitmap[], unsigned int w, unsigned int h, unsigned int Stride)
            {
                unsigned int a, r, g, b, gray, ARGB;
                unsigned int x, y, offset = Stride/4;
                for (y = 0; y < h; ++y) {
                    for (x = 0; x < w; ++x) {
                        ARGB = bitmap[x+(y*offset)];
                        a = ARGB & 0xFF000000;
                        r = (ARGB & 0x00FF0000) >> 16;
                        g = (ARGB & 0x0000FF00) >> 8;
                        b = (ARGB & 0x000000FF);
                        bitmap[x+(y*offset)] = ((255-r) << 16) | ((255-g) << 8) | (255-b) | a;
                    }
                }
                return 0;
            }
        */
        this.InvertColorsMCode := this.MCode((A_PtrSize = 4)
        ? "2,x86:VVdWU4PsCIt8JCiLVCQki0QkIMHvAok8JIXSdF+FwHRbwecCMe2JfCQEi3wkHI00hzH/jXQmAJCLRCQcjQyokIsRg8EEidCJ04Hi/wAA//fQ99OA8v8lAAD/AIHjAP8AAAnYCdCJQfw58XXUg8cBAywkA3QkBDl8JCR1vIPECDHAW15fXcM="
        : "2,x64:VVdWU0SJz0iJy0GJ00SJxsHvAkWFwHRbRTHJRTHAhdJ0UWYPH0QAAESJyQ8fRAAAiciDwQFMjRSDQYsSidCJ1YHi/wAA//fQ99WA8v8lAAD/AIHlAP8AAAnoCdBBiQJBOct1zEGDwAFBAflBAftEOcZ1tTHAW15fXcM=")
    }

    /**
     * Returns an OCR results object for an IRandomAccessStream.
     * Images of other types should be first converted to this format (eg from file, from bitmap).
     * @param RandomAccessStreamOrSoftwareBitmap Pointer or an object containing a ptr to a RandomAccessStream or SoftwareBitmap
     * @param {String} lang OCR language. Default is first from available languages.
     * @param {Integer|Object} transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @param {String} decoder Optional bitmap codec name to decode RandomAccessStream. Default is automatic detection.
     *  Possible values are gif, ico, jpeg, jpegxr, png, tiff, bmp.
     * @returns {OCR.OcrResult} 
     */
    __New(RandomAccessStreamOrSoftwareBitmap, lang := "FirstFromAvailableLanguages", transform := 1, decoder := "") {
        local SoftwareBitmap := 0, RandomAccessStream := 0, width, height, x, y, w, h, __OCR := this.__OCR, scale, grayscale, invertcolors
        __OCR.__ExtractTransformParameters(RandomAccessStreamOrSoftwareBitmap, &transform)
        scale := transform.scale, grayscale := transform.grayscale, invertcolors := transform.invertcolors, rotate := transform.rotate, flip := transform.flip
        __OCR.__ExtractNamedParameters(RandomAccessStreamOrSoftwareBitmap, "x", &x, "y", &y, "w", &w, "h", &h, "lang", &lang, "decoder", &decoder, "RandomAccessStream", &RandomAccessStreamOrSoftwareBitmap, "RAS", &RandomAccessStreamOrSoftwareBitmap, "SoftwareBitmap", &RandomAccessStreamOrSoftwareBitmap)
        __OCR.LoadLanguage(lang)

        try SoftwareBitmap := ComObjQuery(RandomAccessStreamOrSoftwareBitmap, "{689e0708-7eef-483f-963f-da938818e073}") ; ISoftwareBitmap
        if SoftwareBitmap {
            ComCall(8, SoftwareBitmap, "uint*", &width:=0)   ; get_PixelWidth
            ComCall(9, SoftwareBitmap, "uint*", &height:=0)   ; get_PixelHeight
            this.ImageWidth := width, this.ImageHeight := height
            if (Floor(width*scale) > __OCR.MaxImageDimension) or (Floor(height*scale) > __OCR.MaxImageDimension)
               throw ValueError("Image is too big - " width "x" height ".`nIt should be maximum - " __OCR.MaxImageDimension " pixels (with scale applied)")
            if scale != 1 || IsSet(x) || rotate || flip
                SoftwareBitmap := __OCR.TransformSoftwareBitmap(SoftwareBitmap, &width, &height, scale, rotate, flip, x?, y?, w?, h?)
            goto SoftwareBitmapCommon
        }
        RandomAccessStream := RandomAccessStreamOrSoftwareBitmap

        if decoder {
            ComCall(__OCR.Vtbl_GetDecoder.%decoder%, __OCR.BitmapDecoderStatics, "ptr", DecoderGUID:=Buffer(16))
            ComCall(15, __OCR.BitmapDecoderStatics, "ptr", DecoderGUID, "ptr", RandomAccessStream, "ptr*", BitmapDecoder:=ComValue(13,0))   ; CreateAsync
        } else
            ComCall(14, __OCR.BitmapDecoderStatics, "ptr", RandomAccessStream, "ptr*", BitmapDecoder:=ComValue(13,0))   ; CreateAsync
        __OCR.WaitForAsync(&BitmapDecoder)

        BitmapFrame := ComObjQuery(BitmapDecoder, IBitmapFrame := "{72A49A1C-8081-438D-91BC-94ECFC8185C6}")
        ComCall(12, BitmapFrame, "uint*", &width:=0)   ; get_PixelWidth
        ComCall(13, BitmapFrame, "uint*", &height:=0)   ; get_PixelHeight
        if (width > __OCR.MaxImageDimension) or (height > __OCR.MaxImageDimension)
           throw ValueError("Image is too big - " width "x" height ".`nIt should be maximum - " __OCR.MaxImageDimension " pixels")

        BitmapFrameWithSoftwareBitmap := ComObjQuery(BitmapDecoder, IBitmapFrameWithSoftwareBitmap := "{FE287C9A-420C-4963-87AD-691436E08383}")
       if !IsSet(x) && (width < 40 || height < 40 || scale != 1) {
            scale := scale = 1 ? 40.0 / Min(width, height) : scale, this.ImageWidth := Floor(width*scale), this.ImageHeight := Floor(height*scale)
            ComCall(7, __OCR.BitmapTransform, "int", this.ImageWidth) ; put_ScaledWidth
            ComCall(9, __OCR.BitmapTransform, "int", this.ImageHeight) ; put_ScaledHeight
            ComCall(8, BitmapFrame, "uint*", &BitmapPixelFormat:=0) ; get_BitmapPixelFormat
            ComCall(9, BitmapFrame, "uint*", &BitmapAlphaMode:=0) ; get_BitmapAlphaMode
            ComCall(8, BitmapFrameWithSoftwareBitmap, "uint", BitmapPixelFormat, "uint", BitmapAlphaMode, "ptr", __OCR.BitmapTransform, "uint", IgnoreExifOrientation := 0, "uint", DoNotColorManage := 0, "ptr*", SoftwareBitmap:=ComValue(13,0)) ; GetSoftwareBitmapAsync
        } else {
            this.ImageWidth := width, this.ImageHeight := height
            ComCall(6, BitmapFrameWithSoftwareBitmap, "ptr*", SoftwareBitmap:=ComValue(13,0))   ; GetSoftwareBitmapAsync
        }
        __OCR.WaitForAsync(&SoftwareBitmap)
        if IsSet(x) || rotate || flip
            SoftwareBitmap := __OCR.TransformSoftwareBitmap(SoftwareBitmap, &width, &height, scale, rotate, flip, x?, y?, w?, h?)

        SoftwareBitmapCommon:

        if (grayscale || invertcolors || __OCR.DisplayImage) {
            ComCall(15, SoftwareBitmap, "int", 2, "ptr*", BitmapBuffer := ComValue(13,0)) ; LockBuffer
            MemoryBuffer := ComObjQuery(BitmapBuffer, "{fbc4dd2a-245b-11e4-af98-689423260cf8}")
            ComCall(6, MemoryBuffer, "ptr*", MemoryBufferReference := ComValue(13,0)) ; CreateReference
            BufferByteAccess := ComObjQuery(MemoryBufferReference, "{5b0d3235-4dba-4d44-865e-8f1d0e4fd04d}")
            ComCall(3, BufferByteAccess, "ptr*", &SoftwareBitmapByteBuffer:=0, "uint*", &BufferSize:=0) ; GetBuffer

            if invertcolors
                DllCall(__OCR.InvertColorsMCode, "ptr", SoftwareBitmapByteBuffer, "uint", width, "uint", height, "uint", (width*4+3) // 4 * 4, "cdecl uint")
            
            if grayscale
                DllCall(__OCR.GrayScaleMCode, "ptr", SoftwareBitmapByteBuffer, "uint", width, "uint", height, "uint", (width*4+3) // 4 * 4, "cdecl uint")
    
            if __OCR.DisplayImage {
                local hdc := DllCall("GetDC", "ptr", 0, "ptr"), bi := Buffer(40, 0), hbm
                NumPut("uint", 40, "int", width, "int", -height, "ushort", 1, "ushort", 32, bi)
                hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", bi, "uint", 0, "ptr*", &ppvBits:=0, "ptr", 0, "uint", 0, "ptr")
                DllCall("ntdll\memcpy", "ptr", ppvBits, "ptr", SoftwareBitmapByteBuffer, "uint", BufferSize, "cdecl")
                __OCR.DisplayHBitmap(hbm)
            }
            
            BufferByteAccess := "", MemoryBufferReference := "", MemoryBuffer := "", BitmapBuffer := "" ; Release in correct order
        }

        ComCall(6, __OCR.OcrEngine, "ptr", SoftwareBitmap, "ptr*", OcrResult:=ComValue(13,0))   ; RecognizeAsync
        __OCR.WaitForAsync(&OcrResult)
        this.ptr := OcrResult.ptr, ObjAddRef(OcrResult.ptr)

        ; Cleanup
        if RandomAccessStream is __OCR.IBase
            __OCR.CloseIClosable(RandomAccessStream)
        if SoftwareBitmap is __OCR.IBase
            __OCR.CloseIClosable(SoftwareBitmap)

        if scale != 1
            __OCR.NormalizeCoordinates(this, scale)
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
     * added by default with OCR.FromWindow) containing a hWnd property, then that window will be activated,
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
     * @returns {OCR.OcrResult}
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
        if this.HasProp("Relative") {
            if this.Relative.HasOwnProp("Client")
                WinGetClientPos(&rX, &rY,,, this.Relative.Client.hWnd), x += rX + this.Relative.Client.x, y += rY + this.Relative.Client.y
            else if this.Relative.HasOwnProp("Window")
                WinGetPos(&rX, &rY,,, this.Relative.Window.hWnd), x += rX + this.Relative.Window.x, y += rY + this.Relative.Window.y
            else if this.Relative.HasOwnProp("Screen")
                x += this.Relative.Screen.X, y += this.Relative.Screen.Y
        }

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
                    if wordCompareFunc(t, splitNeedle[found.Length+1]) {
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
     * @returns {OCR.OcrResult}
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
     * @returns {OCR.OcrResult}
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
     * @param FileName Either full or relative (to A_WorkingDir) path to the file.
     * @param lang OCR language. Default is first from available languages.
     * @param transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @returns {OCR.OcrResult} 
     */
    static FromFile(FileName, lang?, transform:=1) {
        this.__ExtractTransformParameters(FileName, &transform)
        this.__ExtractNamedParameters(FileName, "lang", &lang, "FileName", &FileName)
        if !(fe := FileExist(FileName)) or InStr(fe, "D")
            throw TargetError("File `"" FileName "`" doesn't exist", -1)
        GUID := this.CLSIDFromString(this.IID_IRandomAccessStream)
        DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", FileName, "uint", Read := 0, "ptr", GUID, "ptr*", IRandomAccessStream:=this.IBase())
        return this(IRandomAccessStream, lang?, transform, this.Vtbl_GetDecoder.HasOwnProp(ext := StrSplit(FileName, ".")[-1]) ? ext : "")
    }

    /**
     * Returns an array of OCR results objects for a PDF file. Locations of the words will be relative to
     * the top left corner of the PDF page.
     * @param FileName Either full or relative (to A_WorkingDir) path to the file.
     * @param lang OCR language. Default is first from available languages.
     * @param transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @param start Page number to start from. Default is first page.
     * @param end Page number to end with (included). Default is last page.
     * @returns {OCR.OcrResult} 
     */
    static FromPDF(FileName, lang?, transform:=1, start:=1, end?) {
        this.__ExtractTransformParameters(FileName, &transform)
        this.__ExtractNamedParameters(FileName, "lang", &lang, "start", &start, "end", &end, "FileName", &FileName)
        if !(fe := FileExist(FileName)) or InStr(fe, "D")
            throw TargetError("File `"" FileName "`" doesn't exist", -1)

        DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", FileName, "uint", Read := 0, "ptr", GUID := this.CLSIDFromString(this.IID_IRandomAccessStream), "ptr*", IRandomAccessStream:=ComValue(13,0))
        PdfDocumentStatics := this.CreateClass("Windows.Data.Pdf.PdfDocument", this.IID_IPdfDocumentStatics) ; If this line is removed then the script throws an error for a second on script exist
        ComCall(8, PdfDocumentStatics, "ptr", IRandomAccessStream, "ptr*", PdfDocument:=this.IBase()) ; LoadFromStreamAsync
        this.WaitForAsync(&PdfDocument)
        this.CloseIClosable(IRandomAccessStream)
        if !IsSet(end) {
            ComCall(7, PdfDocument, "uint*", &end:=0) ; GetPageCount
            if !end
                throw Error("Unable to get PDF page count", -1)
        }
        local results := []
        Loop (end+1-start)
            results.Push(this.FromPDFPage(PdfDocument, start+(A_Index-1), lang?, transform))
        return results
    }

    /**
     * Returns an OCR result object for a PDF page. Locations of the words will be relative to
     * the top left corner of the PDF page.
     * @param FileName Either full or relative (to A_WorkingDir) path to the file.
     * @param Page The page number to OCR. Default is 1.
     * @param lang OCR language. Default is first from available languages.
     * @param transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @returns {OCR.OcrResult} 
     */
    static FromPDFPage(FileName, page:=1, lang?, transform:=1) {
        this.__ExtractTransformParameters(FileName, &transform)
        this.__ExtractNamedParameters(FileName, "page", page, "lang", &lang, "FileName", &FileName)
        if FileName is String {
            if !(fe := FileExist(FileName)) or InStr(fe, "D")
                throw TargetError("File `"" FileName "`" doesn't exist", -1)
            GUID := OCR.CLSIDFromString(OCR.IID_IRandomAccessStream)
            DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", FileName, "uint", Read := 0, "ptr", GUID, "ptr*", IRandomAccessStream:=OCR.IBase())
            PdfDocumentStatics := this.CreateClass("Windows.Data.Pdf.PdfDocument", this.IID_IPdfDocumentStatics)
            ComCall(8, PdfDocumentStatics, "ptr", IRandomAccessStream, "ptr*", PdfDocument:=this.IBase()) ; LoadFromStreamAsync
            this.WaitForAsync(&PdfDocument)
        } else
            PdfDocument := FileName
        ComCall(6, PdfDocument, "uint", page-1, "ptr*", PdfPage:=this.IBase()) ; GetPage
        InMemoryRandomAccessStream := this.CreateClass("Windows.Storage.Streams.InMemoryRandomAccessStream")
        ComCall(6, PdfPage, "ptr", InMemoryRandomAccessStream, "ptr*", asyncInfo:=this.IBase())   ; RenderToStreamAsync
        this.WaitForAsync(&asyncInfo)
        if FileName is String
            this.CloseIClosable(IRandomAccessStream)
        PdfPage := "", PdfDocument := "", IRandomAccessStream := ""
        return this(InMemoryRandomAccessStream, lang?, transform)    
    }

    /**
     * Returns an OCR results object for a given window. Locations of the words will be relative to the
     * window or client area, so for interactions use CoordMode "Window" or "Client". If onlyClientArea
     * contained relative coordinates then Result coordinates will also be relative to the captured area.
     * In that case offsets for Window/Client area are stored in Result.Relative.Client.x and y or .Window.x and y.
     * Additionally, Result.Relative.Screen.x and y are also stored. 
     * @param WinTitle A window title or other criteria identifying the target window.
     * @param lang OCR language. Default is first from available languages.
     * @param transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @param {Number, Object} onlyClientArea Whether only the client area or the whole window should be OCR-d. Default is 0.
     *     This can also be an object which can contain {X,Y,W,H, onlyClientArea} (relative coordinates from where to OCR).
     * @param {Number} mode Different methods of capturing the window. 
     * 0 = uses GetDC with BitBlt
     * 1 = same as 0 but window transparency is turned off beforehand with WinSetTransparent
     * 2 = uses PrintWindow. 
     * 3 = same as 1 but window transparency is turned off beforehand with WinSetTransparent
     * 4 = uses PrintWindow with undocumented PW_RENDERFULLCONTENT flag, allowing capture of hardware-accelerated windows
     * 5 = uses Direct3D11 from UWP Windows.Graphics.Capture (slowest option, but may work with games) 
     *      This may draw a yellow border around the target window in older Windows versions.
     * @returns {OCR.OcrResult} 
     */
    static FromWindow(WinTitle:="", lang?, transform:=1, onlyClientArea:=0, mode:=4) {
        this.__ExtractTransformParameters(WinTitle, &transform)
        local result, X := 0, Y := 0, W := 0, H := 0, sX, sY, hBitMap, hwnd, customRect := 0, scale := transform.scale
        this.__ExtractNamedParameters(WinTitle, "x", &x, "y", &y, "w", &w, "h", &h, "onlyClientArea", &onlyClientArea, "mode", &mode, "lang", &lang, "WinTitle", &Wintitle)
        this.__ExtractNamedParameters(onlyClientArea, "x", &x, "y", &y, "w", &w, "h", &h, "onlyClientArea", &onlyClientArea)
        if (x !=0 || y != 0 || w != 0 || h != 0)
            customRect := 1
        if IsObject(WinTitle)
            WinTitle := ""
        if !(hWnd := WinExist(WinTitle))
            throw TargetError("Target window not found", -1)
        if DllCall("IsIconic", "uptr", hwnd)
            DllCall("ShowWindow", "uptr", hwnd, "int", 4)
        if mode < 4 && mode&1 {
            oldStyle := WinGetExStyle(hwnd), i := 0
            WinSetTransparent(255, hwnd)
            While (WinGetTransparent(hwnd) != 255 && ++i < 30)
                Sleep 100
        }

        WinGetPos(&wX, &wY, &wW, &wH, hWnd)
        If onlyClientArea = 1 {
            WinGetClientPos(&cX, &cY, &cW, &cH, hWnd)
            W := W || cW, H := H || cH, sX := X + cX, sY := Y + cY  ; Calculate final X and Y screen coordinates
        } else {
            W := W || wW, H := H || wH, sX := X + wX, sY := Y + wY
        }

        if mode = 5 {
            /*
                If we are capturing the whole window, then WinGetPos/MouseGetPos might include hidden borders.
                Eg (0,0) might be (-11, -11) for Direct3D, meaning (11,11) by WinGetPos is (0,0) for Direct3D.
                These offsets are calculated and stored in offsetX, offsetY, and if only the window
                area is captured then the result object coordinates are adjusted accordingly.

                If the SoftwareBitmap needs to be transformed in any way (eg scale or custom rect is
                provided) then we need to offset coordinates and possibly width/height as well.

            */
            SoftwareBitmap := this.CreateDirect3DSoftwareBitmapFromWindow(hWnd)

            local offsetX := 0, offsetY := 0, sbW := SoftwareBitmap.W, sbH := SoftwareBitmap.H, sbX := SoftwareBitmap.X, sbY := SoftwareBitmap.Y

            if scale != 1 || transform.rotate || transform.flip || customRect || onlyClientArea {
                ; The bounds need to fit inside the SoftwareBitmap bounds, so possibly X,Y need to be adjusted along with W,H
                local tX := X, tY := Y, tW := W, tH := H
                if onlyClientArea
                    tX -= SoftwareBitmap.X-cX, tY -= SoftwareBitmap.Y-cY
                else
                    tX -= SoftwareBitmap.X-wX, tY -= SoftwareBitmap.Y-wY
                if tX < 0 ; If resulting coordinates are negative then adjust width and height accordingly 
                    tW += tX, offsetX := -tX, tX := 0
                if tY < 0
                    tH += tY, offsetY := -tY, tY := 0
                tW := Min(sbW-tX, tW), tH := Min(sbH-tY, tH)

                SoftwareBitmap := this.TransformSoftwareBitmap(SoftwareBitmap, &sbW, &sbH, scale, transform.rotate, transform.flip, tX, tY, tW, tH)
                transform.scale := 1, transform.rotate := 0, transform.flip := 0
            }
            result := this(SoftwareBitmap, lang?, transform)
        } else {
            hBitMap := this.CreateHBitmap(X, Y, W, H, {hWnd:hWnd, onlyClientArea:onlyClientArea, mode:(mode//2)}, scale)
            if mode&1
                WinSetExStyle(oldStyle, hwnd)
            result := this(this.HBitmapToSoftwareBitmap(hBitMap,, transform), lang?)
        }

        result.Relative := {Screen:{X:sX, Y:sY, W:W, H:H}}
        , result.Relative.%(onlyClientArea = 1 ? "Client" : "Window")% := {X:X, Y:Y, W:W, H:H, hWnd:hWnd}
        this.NormalizeCoordinates(result, scale)
        if mode = 5 && !onlyClientArea
            result.OffsetCoordinates(offsetX, offsetY)
        return result
    }

    /**
     * Returns an OCR results object for the whole desktop. Locations of the words will be relative to
     * the screen (CoordMode "Screen") in a single-monitor setup. If "monitor" argument is specified
     * then coordinates might be relative to the monitor, whereas relative offsets will be stored in
     * Result.Relative.Screen.x and y properties. 
     * @param lang OCR language. Default is first from available languages.
     * @param transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @param monitor The monitor from which to get the desktop area. Default is primary monitor.
     *   If screen scaling between monitors differs, then use DllCall("SetThreadDpiAwarenessContext", "ptr", -3)
     * @returns {OCR.OcrResult} 
     */
    static FromDesktop(lang?, transform:=1, monitor?) {
        if IsSet(lang) {
            this.__ExtractTransformParameters(lang, &transform)
            lang := lang.HasProp("lang") ? lang : unset
        }
        MonitorGet(monitor?, &Left, &Top, &Right, &Bottom)
        return this.FromRect(Left, Top, Right-Left, Bottom-Top, lang?, transform)
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
     * @param transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @returns {OCR.OcrResult} 
     */
    static FromRect(x, y?, w?, h?, lang?, transform:=1) {
        this.__ExtractTransformParameters(x, &transform)
        this.__ExtractNamedParameters(x, "y", &y, "w", &w, "h", &h, "lang", &lang, "x", &x)
        local scale := transform.scale
            , hBitmap := this.CreateHBitmap(X, Y, W, H,, scale)
            , result := this(this.HBitmapToSoftwareBitmap(hBitmap,, transform), lang?)
        result.Relative := {Screen:{x:x, y:y, w:w, h:h}}
        return this.NormalizeCoordinates(result, scale)
    }

    /**
     * Returns an OCR results object from a bitmap. Locations of the words will be relative
     * to the top left corner of the bitmap.
     * @param bitmap A pointer to a GDIP Bitmap object, or HBITMAP, or an object with a ptr property
     *  set to one of the two.
     * @param lang OCR language. Default is first from available languages.
     * @param transform Either a scale factor number, or an object {scale:Float, grayscale:Boolean, invertcolors:Boolean, rotate: 0 | 90 | 180 | 270, flip: 0 | "x" | "y"}
     * @param hDC Optional: a device context for the bitmap. If omitted then the screen DC is used.
     * @returns {OCR.OcrResult} 
     */
    static FromBitmap(bitmap, lang?, transform:=1, hDC?) {
        this.__ExtractTransformParameters(bitmap, &transform)
        local result, pDC, hBitmap, hBM2, oBM, oBM2, pBitmapInfo := Buffer(32, 0), W, H, scale := transform.scale
        this.__ExtractNamedParameters(bitmap, "hDC", &hDC, "lang", &lang, "hBitmap", &bitmap, "pBitmap", &bitmap, "bitmap", &bitmap)
        if !DllCall("GetObject", "ptr", bitmap, "int", pBitmapInfo.Size, "ptr", pBitmapInfo) {
            DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "UPtr", bitmap, "UPtr*", &hBitmap:=0, "Int", 0xffffffff)
            DllCall("GetObject", "ptr", hBitmap, "int", pBitmapInfo.Size, "ptr", pBitmapInfo)
        } else
            hBitmap := bitmap

        W := NumGet(pBitmapInfo, 4, "int"), H := NumGet(pBitmapInfo, 8, "int")

        if scale != 1 || (W && H && (W < 40 || H < 40)) {
            sW := Ceil(W * scale), sH := Ceil(H * scale)

            hDC := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
            , oBM := DllCall("SelectObject", "Ptr", hDC, "Ptr", hBitmap, "Ptr")
            , pDC := DllCall("CreateCompatibleDC", "Ptr", hDC, "Ptr")
            , hBM2 := DllCall("CreateCompatibleBitmap", "Ptr", hDC, "Int", Max(40, sW), "Int", Max(40, sH), "Ptr")
            , oBM2 := DllCall("SelectObject", "Ptr", pDC, "Ptr", hBM2, "Ptr")
            if sW < 40 || sH < 40 ; Fills the bitmap so it's at least 40x40, which seems to improve recognition
                DllCall("StretchBlt", "Ptr", pDC, "Int", 0, "Int", 0, "Int", Max(40,sW), "Int", Max(40,sH), "Ptr", hDC, "Int", 0, "Int", 0, "Int", 1, "Int", 1, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY. 
            PrevStretchBltMode := DllCall("SetStretchBltMode", "Ptr", PDC, "Int", 3, "Int") ; COLORONCOLOR
            , DllCall("StretchBlt", "Ptr", pDC, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", hDC, "Int", 0, "Int", 0, "Int", W, "Int", H, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY
            , DllCall("SetStretchBltMode", "Ptr", PDC, "Int", PrevStretchBltMode)
            , DllCall("SelectObject", "Ptr", pDC, "Ptr", oBM2)
            , DllCall("SelectObject", "Ptr", hDC, "Ptr", oBM)
            , DllCall("DeleteDC", "Ptr", hDC)
            result := this(this.HBitmapToSoftwareBitmap(hBM2, pDC, transform), lang?)
            this.NormalizeCoordinates(result, scale)
            DllCall("DeleteDC", "Ptr", pDC)
            , DllCall("DeleteObject", "UPtr", hBM2)
            return result
        } 
        return this(this.HBitmapToSoftwareBitmap(hBitmap, hDC?, transform), lang?)
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
        local hString, Language:=this.IBase(), OcrEngine:=this.IBase()
        if this.HasOwnProp("CurrentLanguage") && this.HasOwnProp("OcrEngine") && this.CurrentLanguage = lang
            return
        if (lang = "FirstFromAvailableLanguages")
            ComCall(10, this.OcrEngineStatics, "ptr*", OcrEngine)   ; TryCreateFromUserProfileLanguages
        else {
            hString := this.CreateHString(lang)
            , ComCall(6, this.LanguageFactory, "ptr", hString, "ptr*", Language)   ; CreateLanguage
            , this.DeleteHString(hString)
            , ComCall(9, this.OcrEngineStatics, "ptr", Language, "ptr*", OcrEngine)   ; TryCreateFromLanguage
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
     * @returns {OCR.OcrResult} 
     */
    static WaitText(needle, timeout:=-1, func?, casesense:=False, comparefunc?) {
        local endTime := A_TickCount+timeout, result, line, total
        if !IsSet(func)
            func := this.FromDesktop
        if !IsSet(comparefunc)
            comparefunc := InStr.Bind(,,casesense)
        While timeout > 0 ? (A_TickCount < endTime) : 1 {
            result := func(), total := ""
            for line in result.Lines
                total .= line.Text "`n"
            if comparefunc(Trim(total, "`n"), needle)
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

    ; Scales and optionally crops a SoftwareBitmap. Crop parameters need to not be scale-adjusted.
    ; Rotation can be clockwise 0, 90, 180, or 270 degrees
    ; Flip: 0 = no flip, 1 = around y-axis, 2 = around x-axis
    static TransformSoftwareBitmap(SoftwareBitmap, &sbW, &sbH, scale:=1, rotate:=0, flip:=0, X?, Y?, W?, H?) {
        InMemoryRandomAccessStream := this.SoftwareBitmapToRandomAccessStream(SoftwareBitmap)

        ComCall(this.Vtbl_GetDecoder.png, this.BitmapDecoderStatics, "ptr", DecoderGUID:=Buffer(16))
        ComCall(15, this.BitmapDecoderStatics, "ptr", DecoderGUID, "ptr", InMemoryRandomAccessStream, "ptr*", BitmapDecoder:=ComValue(13,0))   ; CreateAsync
        this.WaitForAsync(&BitmapDecoder)

        BitmapFrameWithSoftwareBitmap := ComObjQuery(BitmapDecoder, IBitmapFrameWithSoftwareBitmap := "{FE287C9A-420C-4963-87AD-691436E08383}")
        BitmapFrame := ComObjQuery(BitmapDecoder, IBitmapFrame := "{72A49A1C-8081-438D-91BC-94ECFC8185C6}")

        BitmapTransform := this.CreateClass("Windows.Graphics.Imaging.BitmapTransform")

        local sW := Floor(sbW*scale), sH := Floor(sbH*scale), intermediate
        if scale != 1 {
            ; First the bitmap is scaled, then cropped
            ComCall(7, BitmapTransform, "uint", sW) ; put_ScaledWidth
            ComCall(9, BitmapTransform, "uint", sH) ; put_ScaledHeight
        }
        if rotate {
            ComCall(15, BitmapTransform, "uint", rotate//90) ; put_Rotation
            if rotate = 90 || rotate = 270
                intermediate := sW, sW := sH, sH := intermediate
        }
        if flip
            ComCall(13, BitmapTransform, "uint", flip) ; put_Flip

        if IsSet(X) {
            bounds := Buffer(16,0), NumPut("int", Floor(X*scale), "int", Floor(Y*scale), "int", Floor(Min(sbW-X, W)*scale), "int", Floor(Min(sbH-Y, H)*scale), bounds)
            ComCall(17, BitmapTransform, "ptr", bounds) ; put_Bounds
        }
        ComCall(8, BitmapFrame, "uint*", &BitmapPixelFormat:=0) ; get_BitmapPixelFormat
        ComCall(9, BitmapFrame, "uint*", &BitmapAlphaMode:=0) ; get_BitmapAlphaMode
        ComCall(8, BitmapFrameWithSoftwareBitmap, "uint", BitmapPixelFormat, "uint", BitmapAlphaMode, "ptr", BitmapTransform, "uint", IgnoreExifOrientation := 0, "uint", DoNotColorManage := 0, "ptr*", SoftwareBitmap:=ComValue(13,0)) ; GetSoftwareBitmapTransformedAsync

        this.WaitForAsync(&SoftwareBitmap)
        ; this.CloseIClosable(BitmapFrameWithSoftwareBitmap) ; Implemented, but is it necessary?
        this.CloseIClosable(InMemoryRandomAccessStream)
        sbW := sW, sbH := sH
        return SoftwareBitmap
    }

    static CreateDIBSection(w, h, hdc?, bpp:=32, &ppvBits:=0) {
        local hdc2 := IsSet(hdc) ? hdc : DllCall("GetDC", "Ptr", 0, "UPtr")
        , bi := Buffer(40, 0), hbm
        NumPut("int", 40, "int", w, "int", h, "ushort", 1, "ushort", bpp, "int", 0, bi)
        hbm := DllCall("CreateDIBSection", "uint", hdc2, "ptr" , bi, "uint" , 0, "uint*", &ppvBits:=0, "uint" , 0, "uint" , 0)
        if !IsSet(hdc)
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc2)
        return hbm
    }

    /**
     * Creates an hBitmap of a region of the screen or a specific window
     * @param X Captured rectangle X coordinate. This is relative to the screen unless hWnd is specified,
     *  in which case it may be relative to the window/client
     * @param Y Captured rectangle Y coordinate.
     * @param W Captured rectangle width.
     * @param H Captured rectangle height.
     * @param {Integer|Object} hWnd Window handle which to capture. Coordinates will be relative to the window. 
     *  hWnd may also be an object {hWnd, onlyClientArea, mode} where onlyClientArea:1 means the client area will be captured instead of the whole window (and X, Y will also be relative to client)
     *  mode 0 uses GetDC + StretchBlt, mode 1 uses PrintWindow, mode 2 uses PrintWindow with undocumented PW_RENDERFULLCONTENT flag. 
     *  Default is mode 2.
     * @param {Integer} scale 
     * @returns {OCR.IBase} 
     */
    static CreateHBitmap(X, Y, W, H, hWnd:=0, scale:=1) {
        local sW := Ceil(W*scale), sH := Ceil(H*scale), onlyClientArea := 0, mode := 2, HDC, obm, hbm, pdc, hbm2
        if hWnd {
            if IsObject(hWnd)
                onlyClientArea := hWnd.HasOwnProp("onlyClientArea") ? hWnd.onlyClientArea : onlyClientArea, mode := hWnd.HasOwnProp("mode") ? hWnd.mode : mode, hWnd := hWnd.hWnd
            HDC := DllCall("GetDCEx", "Ptr", hWnd, "Ptr", 0, "Int", 2|!onlyClientArea, "Ptr")
            if mode > 0 {
                PDC := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
                HBM := DllCall("CreateCompatibleBitmap", "Ptr", HDC, "Int", Max(40,X+W), "Int", Max(40,Y+H), "Ptr")
                , OBM := DllCall("SelectObject", "Ptr", PDC, "Ptr", HBM, "Ptr")
                , DllCall("PrintWindow", "Ptr", hWnd, "Ptr", PDC, "UInt", (mode=2?2:0)|!!onlyClientArea)
                if scale != 1 || X != 0 || Y != 0 {
                    PDC2 := DllCall("CreateCompatibleDC", "Ptr", PDC, "Ptr")
                    , HBM2 := DllCall("CreateCompatibleBitmap", "Ptr", PDC, "Int", Max(40,sW), "Int", Max(40,sH), "Ptr")
                    , OBM2 := DllCall("SelectObject", "Ptr", PDC2, "Ptr", HBM2, "Ptr")
                    , PrevStretchBltMode := DllCall("SetStretchBltMode", "Ptr", PDC, "Int", 3, "Int") ; COLORONCOLOR
                    , DllCall("StretchBlt", "Ptr", PDC2, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", PDC, "Int", X, "Int", Y, "Int", W, "Int", H, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY
                    , DllCall("SetStretchBltMode", "Ptr", PDC, "Int", PrevStretchBltMode)
                    , DllCall("SelectObject", "Ptr", PDC2, "Ptr", obm2)
                    , DllCall("DeleteDC", "Ptr", PDC)
                    , DllCall("DeleteObject", "UPtr", HBM)
                    , hbm := hbm2, pdc := pdc2
                }
                DllCall("SelectObject", "Ptr", PDC, "Ptr", OBM)
                , DllCall("DeleteDC", "Ptr", HDC)
                , oHBM := this.IBase(HBM), oHBM.DC := PDC
                return oHBM.DefineProp("__Delete", {call:(this, *)=>(DllCall("DeleteObject", "Ptr", this), DllCall("DeleteDC", "Ptr", this.DC))})
            }
        } else {
            HDC := DllCall("GetDC", "Ptr", 0, "Ptr")
        }
        PDC := DllCall("CreateCompatibleDC", "Ptr", HDC, "Ptr")
        , HBM := DllCall("CreateCompatibleBitmap", "Ptr", HDC, "Int", Max(40,sW), "Int", Max(40,sH), "Ptr")
        , OBM := DllCall("SelectObject", "Ptr", PDC, "Ptr", HBM, "Ptr")
        if sW < 40 || sH < 40 ; Fills the bitmap so it's at least 40x40, which seems to improve recognition
            DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", Max(40,sW), "Int", Max(40,sH), "Ptr", HDC, "Int", X, "Int", Y, "Int", 1, "Int", 1, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY. 
        PrevStretchBltMode := DllCall("SetStretchBltMode", "Ptr", PDC, "Int", 3, "Int") ; COLORONCOLOR
        , DllCall("StretchBlt", "Ptr", PDC, "Int", 0, "Int", 0, "Int", sW, "Int", sH, "Ptr", HDC, "Int", X, "Int", Y, "Int", W, "Int", H, "UInt", 0x00CC0020 | this.CAPTUREBLT) ; SRCCOPY
        , DllCall("SetStretchBltMode", "Ptr", PDC, "Int", PrevStretchBltMode)
        , DllCall("SelectObject", "Ptr", PDC, "Ptr", OBM)
        , DllCall("DeleteDC", "Ptr", HDC)
        , oHBM := this.IBase(HBM), oHBM.DC := PDC
        return oHBM.DefineProp("__Delete", {call:(this, *)=>(DllCall("DeleteObject", "Ptr", this), DllCall("ReleaseDC", "Ptr", 0, "Ptr", this.DC))})
    }

    static CreateDirect3DSoftwareBitmapFromWindow(hWnd) {
        static init := 0, DXGIDevice, Direct3DDevice, Direct3D11CaptureFramePoolStatics, GraphicsCaptureItemInterop, GraphicsCaptureItemGUID, D3D_Device, D3D_Context
        local x, y, w, h, rect
        if !init {
            DllCall("LoadLibrary","str","DXGI")
            DllCall("LoadLibrary","str","D3D11")
            DllCall("LoadLibrary","str","Dwmapi")
            DllCall("D3D11\D3D11CreateDevice", "ptr", 0, "int", D3D_DRIVER_TYPE_HARDWARE := 1, "ptr", 0, "uint", D3D11_CREATE_DEVICE_BGRA_SUPPORT := 0x20, "ptr", 0, "uint", 0, "uint", D3D11_SDK_VERSION := 7, "ptr*", D3D_Device:=ComValue(13, 0), "ptr*", 0, "ptr*", D3D_Context:=ComValue(13, 0))
            DXGIDevice := ComObjQuery(D3D_Device, IID_IDXGIDevice := "{54ec77fa-1377-44e6-8c32-88fd5f44c84c}")
            DllCall("D3D11\CreateDirect3D11DeviceFromDXGIDevice", "ptr", DXGIDevice, "ptr*", GraphicsDevice:=ComValue(13, 0))
            Direct3DDevice := ComObjQuery(GraphicsDevice, IDirect3DDevice := "{A37624AB-8D5F-4650-9D3E-9EAE3D9BC670}")
            Direct3D11CaptureFramePoolStatics := this.CreateClass("Windows.Graphics.Capture.Direct3D11CaptureFramePool", IDirect3D11CaptureFramePoolStatics := "{7784056a-67aa-4d53-ae54-1088d5a8ca21}")
            GraphicsCaptureItemStatics := this.CreateClass("Windows.Graphics.Capture.GraphicsCaptureItem", IGraphicsCaptureItemStatics := "{A87EBEA5-457C-5788-AB47-0CF1D3637E74}")
            GraphicsCaptureItemInterop := ComObjQuery(GraphicsCaptureItemStatics, IGraphicsCaptureItemInterop := "{3628E81B-3CAC-4C60-B7F4-23CE0E0C3356}")
            GraphicsCaptureItemGUID := Buffer(16,0)
            DllCall("ole32\CLSIDFromString", "wstr", IGraphicsCaptureItem := "{79c3f95b-31f7-4ec2-a464-632ef5d30760}", "ptr", GraphicsCaptureItemGUID)
            init := 1
        }
        ; INIT done

        DllCall("Dwmapi.dll\DwmGetWindowAttribute", "ptr", hWnd, "uint", DWMWA_EXTENDED_FRAME_BOUNDS := 9, "ptr", rect := Buffer(16,0), "uint", 16)
        x := NumGet(rect, 0, "int"), y := NumGet(rect, 4, "int"), w := NumGet(rect, 8, "int") - x, h := NumGet(rect, 12, "int") - y
        ComCall(6, Direct3D11CaptureFramePoolStatics, "ptr", Direct3DDevice, "int", B8G8R8A8UIntNormalized := 87, "int", numberOfBuffers := 2, "int64", (h << 32) | w, "ptr*", Direct3D11CaptureFramePool:=ComValue(13, 0))   ; Direct3D11CaptureFramePool.Create
        if ComCall(3, GraphicsCaptureItemInterop, "ptr", hWnd, "ptr", GraphicsCaptureItemGUID, "ptr*", GraphicsCaptureItem:=ComValue(13, 0), "uint") {   ; IGraphicsCaptureItemInterop::CreateForWindow
            this.CloseIClosable(Direct3D11CaptureFramePool)
            throw Error("Failed to capture GraphicsItem of window",, -1)
        }
        ComCall(10, Direct3D11CaptureFramePool, "ptr", GraphicsCaptureItem, "ptr*", GraphicsCaptureSession:=ComValue(13, 0))   ; Direct3D11CaptureFramePool.CreateCaptureSession

        GraphicsCaptureSession2 := ComObjQuery(GraphicsCaptureSession, IGraphicsCaptureSession2 := "{2c39ae40-7d2e-5044-804e-8b6799d4cf9e}")
        ComCall(7, GraphicsCaptureSession2, "int", 0)   ; GraphicsCaptureSession.IsCursorCaptureEnabled put

        if (Integer(StrSplit(A_OSVersion, ".")[3]) >= 20348) { ; hide border
            GraphicsCaptureSession3 := ComObjQuery(GraphicsCaptureSession, IGraphicsCaptureSession3 := "{f2cdd966-22ae-5ea1-9596-3a289344c3be}")
            ComCall(7, GraphicsCaptureSession3, "int", 0)   ; GraphicsCaptureSession.IsBorderRequired put
        }
        ComCall(6, GraphicsCaptureSession)   ; GraphicsCaptureSession.StartCapture
        Loop {
            ComCall(7, Direct3D11CaptureFramePool, "ptr*", Direct3D11CaptureFrame:=ComValue(13, 0))   ; Direct3D11CaptureFramePool.TryGetNextFrame
            if (Direct3D11CaptureFrame.ptr != 0)
                break
        }
        ComCall(6, Direct3D11CaptureFrame, "ptr*", Direct3DSurface:=ComValue(13, 0))   ; Direct3D11CaptureFrame.Surface

        ComCall(11, this.SoftwareBitmapStatics, "ptr", Direct3DSurface, "ptr*", SoftwareBitmap:=ComValue(13, 0)) ; SoftwareBitmap::CreateCopyFromSurfaceAsync
        OCR.WaitForAsync(&SoftwareBitmap)

        this.CloseIClosable(Direct3D11CaptureFramePool)
        this.CloseIClosable(GraphicsCaptureSession)
        if GraphicsCaptureSession2 {
            this.CloseIClosable(GraphicsCaptureSession2)
        }
        if IsSet(GraphicsCaptureSession3) {
            this.CloseIClosable(GraphicsCaptureSession3)
        }
        this.CloseIClosable(Direct3D11CaptureFrame)
        this.CloseIClosable(Direct3DSurface)

        SoftwareBitmap.x := x, SoftwareBitmap.y := y, SoftwareBitmap.w := w, SoftwareBitmap.h := h
        return SoftwareBitmap
    }

    static HBitmapToRandomAccessStream(hBitmap) {
        static PICTYPE_BITMAP := 1
             , BSOS_DEFAULT   := 0
             , sz := 8 + A_PtrSize*2
        local PICTDESC, riid, size, pIRandomAccessStream
             
        DllCall("Ole32\CreateStreamOnHGlobal", "Ptr", 0, "UInt", true, "Ptr*", pIStream:=ComValue(13,0), "UInt")
        , PICTDESC := Buffer(sz, 0)
        , NumPut("uint", sz, "uint", PICTYPE_BITMAP, "ptr", IsInteger(hBitmap) ? hBitmap : hBitmap.ptr, PICTDESC)
        , riid := this.CLSIDFromString(this.IID_IPicture)
        , DllCall("OleAut32\OleCreatePictureIndirect", "Ptr", PICTDESC, "Ptr", riid, "UInt", 0, "Ptr*", pIPicture:=ComValue(13,0), "UInt")
        , ComCall(15, pIPicture, "Ptr", pIStream, "UInt", true, "uint*", &size:=0, "UInt") ; IPicture::SaveAsFile
        , riid := this.CLSIDFromString(this.IID_IRandomAccessStream)
        , DllCall("ShCore\CreateRandomAccessStreamOverStream", "Ptr", pIStream, "UInt", BSOS_DEFAULT, "Ptr", riid, "Ptr*", pIRandomAccessStream:=this.IBase(), "UInt")
        Return pIRandomAccessStream
    }

    ; Converts HBITMAP to SoftwareBitmap. NOTE: SetStretchBltMode HALFTONE breaks this
    ; The optional transform parameter may contain {grayscale, invertcolors}
    static HBitmapToSoftwareBitmap(hBitmap, hDC?, transform?) {
        local bi := Buffer(40, 0), W, H, BitmapBuffer, MemoryBuffer, MemoryBufferReference, BufferByteAccess, BufferSize
        hDC := (hBitmap is OCR.IBase ? hBitmap.DC : (hDC ?? dhDC := DllCall("GetDC", "Ptr", 0, "UPtr")))

        NumPut("uint", 40, bi, 0)
        DllCall("GetDIBits", "ptr", hDC, "ptr", hBitmap, "uint", 0, "uint", 0, "ptr", 0, "ptr", bi, "uint", 0)
        W := NumGet(bi, 4, "int"), H := NumGet(bi, 8, "int")

        ComCall(7, this.SoftwareBitmapFactory, "int", 87, "int", W, "int", H, "int", 0, "ptr*", SoftwareBitmap := ComValue(13,0)) ; CreateWithAlpha: Bgra8 & Premultiplied
        ComCall(15, SoftwareBitmap, "int", 2, "ptr*", BitmapBuffer := ComValue(13,0)) ; LockBuffer
        MemoryBuffer := ComObjQuery(BitmapBuffer, "{fbc4dd2a-245b-11e4-af98-689423260cf8}")
        ComCall(6, MemoryBuffer, "ptr*", MemoryBufferReference := ComValue(13,0)) ; CreateReference
        BufferByteAccess := ComObjQuery(MemoryBufferReference, "{5b0d3235-4dba-4d44-865e-8f1d0e4fd04d}")
        ComCall(3, BufferByteAccess, "ptr*", &SoftwareBitmapByteBuffer:=0, "uint*", &BufferSize:=0) ; GetBuffer

        NumPut("short", 32, "short", 0, bi, 14), NumPut("int", -H, bi, 8) ; Negative height to get correctly oriented image
        DllCall("GetDIBits", "ptr", hDC, "ptr", hBitmap, "uint", 0, "uint", H, "ptr", SoftwareBitmapByteBuffer, "ptr", bi, "uint", 0)
        if IsSet(transform) {
            if (transform.HasProp("grayscale") && transform.grayscale)
                DllCall(this.GrayScaleMCode, "ptr", SoftwareBitmapByteBuffer, "uint", w, "uint", h, "uint", (w*4+3) // 4 * 4, "cdecl uint")
            if (transform.HasProp("invertcolors") && transform.invertcolors)
                DllCall(this.InvertColorsMCode, "ptr", SoftwareBitmapByteBuffer, "uint", w, "uint", h, "uint", (w*4+3) // 4 * 4, "cdecl uint")
        }
        
        if IsSet(dhDC)
            DllCall("DeleteDC", "ptr", dhDC)
        BufferByteAccess := "", MemoryBufferReference := "", MemoryBuffer := "", BitmapBuffer := "" ; Release in correct order

        return SoftwareBitmap
    }

    static MCode(mcode) {
        static e := Map('1', 4, '2', 1), c := (A_PtrSize=8) ? "x64" : "x86"
        if (!regexmatch(mcode, "^([0-9]+),(" c ":|.*?," c ":)([^,]+)", &m))
          return
        if (!DllCall("crypt32\CryptStringToBinary", "str", m.3, "uint", 0, "uint", e[m.1], "ptr", 0, "uint*", &s := 0, "ptr", 0, "ptr", 0))
          return
        p := DllCall("GlobalAlloc", "uint", 0, "ptr", s, "ptr")
        if (c="x64")
          DllCall("VirtualProtect", "ptr", p, "ptr", s, "uint", 0x40, "uint*", &op := 0)
        if (DllCall("crypt32\CryptStringToBinary", "str", m.3, "uint", 0, "uint", e[m.1], "ptr", p, "uint*", &s, "ptr", 0, "ptr", 0))
          return p
        DllCall("GlobalFree", "ptr", p)
      }

    static DisplayHBitmap(hBitmap) {
        local gImage := Gui("-DPIScale"), W, H
        , hPic := gImage.Add("Text", "0xE w640 h640")
        SendMessage(0x172, 0, hBitmap,, hPic.hWnd)
        hPic.GetPos(,,&W, &H)
        gImage.Show("w" (W+20) " H" (H+20))
        WinWaitClose gImage
    }

    static SoftwareBitmapToRandomAccessStream(SoftwareBitmap) {
        InMemoryRandomAccessStream := this.CreateClass("Windows.Storage.Streams.InMemoryRandomAccessStream")
        ComCall(8, this.BitmapEncoderStatics, "ptr", encoderId := Buffer(16, 0)) ; IBitmapEncoderStatics::PngEncoderId
        ComCall(13, this.BitmapEncoderStatics, "ptr", encoderId, "ptr", InMemoryRandomAccessStream, "ptr*", BitmapEncoder:=ComValue(13,0)) ; IBitmapEncoderStatics::CreateAsync
        this.WaitForAsync(&BitmapEncoder)
        BitmapEncoderWithSoftwareBitmap := ComObjQuery(BitmapEncoder, "{686cd241-4330-4c77-ace4-0334968b1768}")
        ComCall(6, BitmapEncoderWithSoftwareBitmap, "ptr", SoftwareBitmap) ; SetSoftwareBitmap
        ComCall(19, BitmapEncoder, "ptr*", asyncAction:=ComValue(13,0)) ; FlushAsync
        this.WaitForAsync(&asyncAction)
        ComCall(11, InMemoryRandomAccessStream, "int64", 0) ; Seek to beginning
        return InMemoryRandomAccessStream
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
            Sleep this.PerformanceMode ? -1 : 1
        }
        ComCall(8, obj, "ptr*", ObjectResult:=this.IBase())   ; GetResults
        obj := ObjectResult
    }

    static CloseIClosable(pClosable) {
        static IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}"
        local Close := ComObjQuery(pClosable, IClosable)
        ComCall(6, Close)   ; Close
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

    static __ExtractNamedParameters(obj, params*) {
        local i := 0
        if !IsObject(obj) || Type(obj) != "Object"
            return 0
        Loop params.Length // 2 {
            name := params[++i], value := params[++i]
            if obj.HasProp(name)
                %value% := obj.%name%
        }
        return 1
    }

    static __ExtractTransformParameters(obj, &transform) {
        local scale := 1, grayscale := 0, invertcolors := 0, rotate := 0, flip := 0
        if IsObject(obj)
            this.__ExtractNamedParameters(obj, "scale", &scale, "grayscale", &grayscale, "invertcolors", &invertcolors, "rotate", &rotate, "flip", &flip, "transform", &transform)

        if IsObject(transform) {
            for prop in ["scale", "grayscale", "invertcolors", "rotate", "flip"]
                if !transform.HasProp(prop)
                    transform.%prop% := %prop%
        } else
            transform := {scale:scale, grayscale:grayscale, invertcolors:invertcolors, rotate:rotate, flip:flip}
    
        transform.flip := transform.flip = "y" ? 1 : transform.flip = "x" ? 2 : transform.flip
    }

    OffsetCoordinates(offsetX?, offsetY?) {
        if !IsSet(offsetX) || !IsSet(offsetY) {
            if this.HasOwnProp("Relative") {
                if this.Relative.HasOwnProp("Client")
                    offsetX := this.Relative.Client.x, offsetY := this.Relative.Client.x
                else if this.Relative.HasOwnProp("Window")
                    offsetX := this.Relative.Window.x, offsetY := this.Relative.Window.y
                else
                    throw Error("No appropriate Relative property found",, -1)
            } else
                throw Error("No Relative property found",, -1)
        }
        if offsetX = 0 && offsetY = 0
            return this
        local word
        for word in this.Words
            word.x += offsetX, word.y += offsetY, word.BoundingRect := {X:word.x, Y:word.y, W:word.w, H:word.h}
        return this
    }

    /**
     * Converts coordinates between screen, window and client.
     * @param X X-coordinate to convert
     * @param Y Y-coordinate to convert
     * @param outX Variable where to store the converted X-coordinate
     * @param outY Variable where to store the converted Y-coordinate
     * @param relativeFrom CoordMode where to convert from. Default is A_CoordModeMouse.
     * @param relativeTo CoordMode where to convert to. Default is Screen.
     * @param winTitle A window title or other criteria identifying the target window. 
     * @param winText If present, this parameter must be a substring from a single text element of the target window.
     * @param excludeTitle Windows whose titles include this value will not be considered.
     * @param excludeText Windows whose text include this value will not be considered.
     */
    static ConvertWinPos(X, Y, &outX, &outY, relativeFrom:="", relativeTo:="screen", winTitle?, winText?, excludeTitle?, excludeText?) {
        relativeFrom := relativeFrom || A_CoordModeMouse
        if relativeFrom = relativeTo {
            outX := X, outY := Y
            return
        }
        local hWnd := WinExist(winTitle?, winText?, excludeTitle?, excludeText?)

        switch relativeFrom, 0 {
            case "screen", "s":
                if relativeTo = "window" || relativeTo = "w" {
                    DllCall("user32\GetWindowRect", "Int", hWnd, "Ptr", RECT := Buffer(16))
                    outX := X-NumGet(RECT, 0, "Int"), outY := Y-NumGet(RECT, 4, "Int")
                } else { 
                    ; screen to client
                    pt := Buffer(8), NumPut("int",X,pt), NumPut("int",Y,pt,4)
                    DllCall("ScreenToClient", "Int", hWnd, "Ptr", pt)
                    outX := NumGet(pt,0,"int"), outY := NumGet(pt,4,"int")
                }
            case "window", "w":
                ; window to screen
                WinGetPos(&outX, &outY,,,hWnd)
                outX += X, outY += Y
                if relativeTo = "client" || relativeTo = "c" {
                    ; screen to client
                    pt := Buffer(8), NumPut("int",outX,pt), NumPut("int",outY,pt,4)
                    DllCall("ScreenToClient", "Int", hWnd, "Ptr", pt)
                    outX := NumGet(pt,0,"int"), outY := NumGet(pt,4,"int")
                }
            case "client", "c":
                ; client to screen
                pt := Buffer(8), NumPut("int",X,pt), NumPut("int",Y,pt,4)
                DllCall("ClientToScreen", "Int", hWnd, "Ptr", pt)
                outX := NumGet(pt,0,"int"), outY := NumGet(pt,4,"int")
                if relativeTo = "window" || relativeTo = "w" { ; screen to window
                    DllCall("user32\GetWindowRect", "Int", hWnd, "Ptr", RECT := Buffer(16))
                    outX -= NumGet(RECT, 0, "Int"), outY -= NumGet(RECT, 4, "Int")
                }
        }
    }
}