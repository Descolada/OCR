# NOTICE
January 2025 version 2 (alpha stage) of this library was published, which introduces multiple breaking changes. Major differences include:
1. Options common to the different OCR functions (such as `scale`, `lang` etc) are now gathered under the `Options` argument
2. OCR.Result objects now contain common methods to all result types (Result, Line, Word, etc) such as `Result.Highlight` and `Result.Click`
3. OCR.FromWindow uses CoordMode from A_CoordModePixel and the option `onlyClientArea` is no longer valid. (applied in 18.02.2025 update)

Since v2 is still in alpha stage, breaking changes are still allowed. If you have any suggestions about the syntax or feature requests, please open an Issue here in GitHub.

# OCR
UWP OCR for AHK v2:
A wrapper for the the UWP Windows.Media.Ocr library. 
Introduced in Windows 10.0.10240.0, this OCR library is included with Windows, no special installs or executables required. Though it might be necessary to install language packs for the OCR, which requires admin access.

Special thanks to AHK forums user malcev, whose OCR function this library is based on.

Examples are included in the Examples folder.

# Table of contents
```
OCR library: a wrapper for the the UWP Windows.Media.Ocr library.
Based on the UWP OCR function for AHK v1 by malcev.

Ways of initiating OCR:
OCR(RandomAccessStreamOrSoftwareBitmap, Options?)
OCR.FromDesktop(Options?, Monitor?)
OCR.FromRect(X, Y, W, H, Options?)
OCR.FromWindow(WinTitle:="", Options?, WinText:="", ExcludeTitle:="", ExcludeText:="")
     Note: the result object coordinates will be in CoordMode "Pixel"
OCR.FromFile(FileName, Options?)
OCR.FromBitmap(bitmap, Options?, hDC?)
OCR.FromPDF(FileName, Options?, Start:=1, End?, Password:="") => returns an array of results for each PDF page
OCR.FromPDFPage(FileName, page, Options?)
  Helper functions for PDF OCR:
     OCR.GetPdfPageCount(FileName, Password:="")
     OCR.GetPdfPageProperties(FileName, Page, Password:="")

Options can be an object containing none or all of these elements:
{
     lang: OCR language. Default is first from available languages.
     scale: a Float scale factor to zoom the image in or out, which might improve detection. 
            The resulting coordinates will be adjusted to scale. Default is 1.
     grayscale: Boolean 0 | 1 whether to convert the image to black-and-white. Default is 0.
     monochrome: 0-255, converts all pixels with luminosity less than the threshold to black, otherwise to white. Default is 0 (no conversion).
     invertcolors: Boolean 0 | 1, whether to invert the colors of the image. Default is 0.
     rotate: 0 | 90 | 180 | 270, can be used to rotate the image clock-wise by degrees. Default is 0.
     flip: 0 | "x" | "y", can be used to flip the image on the x- or y-axis. Default is 0.
     x, y, w, h: can be used to crop the image. This is applied before scaling. Default is no cropping.
     decoder: gif | ico | jpeg | jpegxr | png | tiff | bmp. Optional bitmap codec name to decode RandomAccessStream. Default is automatic detection. 
}

Note: Options also accepts any optional parameters after it like named parameters.
Eg. OCR.FromDesktop({lang:"en-us", monitor:2})

Additional methods:
OCR.GetAvailableLanguages()
OCR.LoadLanguage(lang:="FirstFromAvailableLanguages")
OCR.WaitText(needle, timeout:=-1, func?, casesense:=False, comparefunc?)
     Calls a func (the provided OCR method) until a string is found
OCR.WordsBoundingRect(words*)
     Returns the bounding rectangle for multiple words
OCR.ClearAllHighlights()
     Removes all highlights created by Result.Highlight
OCR.Cluster(objs, eps_x:=-1, eps_y:=-1, minPts:=1, compareFunc?, &noise?)
     Clusters objects (by default based on distance from eachother). Can be used to create more
     accurate "Line" results.
OCR.SortArray(arr, optionsOrCallback:="N", key?)
     Sorts an array in-place, optionally by object keys or using a callback function.
OCR.ReverseArray(arr)
     Reverses an array in-place.
OCR.UniqueArray(arr)
     Returns an array with unique values.
OCR.FlattenArray(arr)
     Returns a one-dimensional array from a multi-dimensional array

Properties:
OCR.MaxImageDimension
MinImageDimension is not documented, but appears to be 40 pixels (source: user FanaticGuru in AutoHotkey forums)
OCR.PerformanceMode
     Increases speed of OCR acquisition by about 20-50ms if set to 1, but also increases CPU usage. Default is 0.
OCR.DisplayImage
     If set to True then the captured image is displayed on the screen before proceeding to OCR-ing the image.

OCR returns an OCR.Result object:
Result.Text         => All recognized text
Result.TextAngle    => Clockwise rotation of the recognized text 
Result.Lines        => Array of all OCR.Line objects
Result.Words        => Array of all OCR.Word objects
Result.ImageWidth   => Used image width
Result.ImageHeight  => Used image height

Result.FindString(Needle, Options?)
     Finds a string in the result. Possible options (see descriptions at the function definition):
     {CaseSense: False, IgnoreLinebreaks: False, AllowOverlap: False, i: 1, x, y, w, h, SearchFunc}
Result.FindStrings(Needle, Options?)
     Finds all strings in the result. 
Result.Filter(callback)
     Returns a filtered result object that contains only words that satisfy the callback function
Result.Crop(x1, y1, x2, y2)
     Crops the result object to contain only results from an area defined by points (x1,y1) and (x2,y2). 

OCR.Line object:
Line.Text         => Recognized text of the line
Line.Words        => Array of Word objects for the Line
Line.x,y,w,h      => Size and location of the Line. 

OCR.Word object:
Word.Text         => Recognized text of the word
Word.x,y,w,h      => Size and location of the Word. 
Word.BoundingRect => Bounding rectangle of the Word in format {x,y,w,h}. 

OCR.Result, OCR.Line, and OCR.Word also all have some common methods:

Result.Click(WhichButton?, ClickCount?, DownOrUp?)
     Clicks an object (Word, FindString result etc)
Result.ControlClick(WinTitle?, WinText?, WhichButton?, ClickCount?, Options?, ExcludeTitle?, ExcludeText?)
     ControlClicks an object (Word, FindString result etc)
Result.Highlight(showTime?, color:="Red", d:=2)
     Highlights a Word, Line, or object with {x,y,w,h} properties on the screen (default: 2 seconds), or removes the highlighting

Additional notes:
Languages are recognized in BCP-47 language tags. Eg. OCR.FromFile("myfile.bmp", {lang: "en-AU"})
Languages can be installed for example with PowerShell (run as admin): Install-Language <language-tag>
     or from Language settings in Settings.
Not all language packs support OCR though. A list of supported language can be gotten from 
Powershell (run as admin) with the following command: Get-WindowsCapability -Online | Where-Object { $_.Name -Like 'Language.OCR*' } 
```

If you wish to support me in this and other projects:
[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/descolada)