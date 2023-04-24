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
OCR(IRandomAccessStream, lang?)
OCR.FromDesktop(lang?, scale:=1)
OCR.FromRect(X, Y, W, H, lang?, scale:=1)
OCR.FromWindow(WinTitle?, lang?, scale:=1, onlyClientArea:=0, mode:=2)
OCR.FromFile(FileName, lang?)
OCR.FromBitmap(HBitmap, lang?)

Additional methods:
OCR.GetAvailableLanguages()
OCR.LoadLanguage(lang:="FirstFromAvailableLanguages")
OCR.WaitText(needle, timeout:=-1, func?, casesense:=False, comparefunc?)
     Calls a func (the provided OCR method) until a string is found
OCR.WordsBoundingRect(words*)
     Returns the bounding rectangle for multiple words


OCR returns an OCR results object:
Result.Text         => All recognized text
Result.TextAngle    => Clockwise rotation of the recognized text 
Result.Lines        => Array of all Line objects
Result.Words        => Array of all Word objects
Result.ImageWidth   => Used image width
Result.ImageHeight  => Used image height

Result.FindString(needle, i:=1, casesense:=False, wordCompareFunc?)
     Finds a string in the result
Result.Click(Obj, WhichButton?, ClickCount?, DownOrUp?)
     Clicks an object (Word, FindString result etc)
Result.ControlClick(obj, WinTitle?, WinText?, WhichButton?, ClickCount?, Options?, ExcludeTitle?, ExcludeText?)
     ControlClicks an object (Word, FindString result etc)
Result.Highlight(obj?, showTime:=2000, color:="Red", d:=2)
     Highlights an object on the screen, or removes the highlighting


Line object:
Line.Text         => Recognized text of the line
Line.Words        => Array of Word objects for the Line

Word object:
Line.Text         => Recognized text of the word
Line.x,y,w,h      => Size and location of the Word. Coordinates are relative to the original image.
Line.BoundingRect => Bounding rectangle of the Word in format {x,y,w,h}. Coordinates are relative to the original image.

Additional notes:
Languages are recognized in BCP-47 language tags. Eg. OCR.FromFile("myfile.bmp", "en-AU")
Languages can be installed for example with PowerShell (run as admin): Install-Language <language-tag>
     or from Language settings in Settings.
Not all language packs support OCR though. A list of supported language can be gotten from 
Powershell (run as admin) with the following command: Get-WindowsCapability -Online | Where-Object { $_.Name -Like 'Language.OCR*' } 
```

If you wish to support me in this and other projects:
[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/descolada)