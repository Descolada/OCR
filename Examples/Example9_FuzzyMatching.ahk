#Requires AutoHotkey v2 
#include ..\Lib\OCR.ahk

; Levensthein distance means how many letter changes need to be made to get from one word to another.
; For example, to get from "AutoHtokey" to "AutoHotkey" we would need to change "t" -> "o" ("AutoHookey")
; and then "o" -> "t", a total of 2 changes. 

; The following example attempts to find words that are less than 3 changes away from "AutoHtokey".
; The search is case-sensitive by default.

result := OCR.FromDesktop(, 2)
for word in result.Words {
    if LD(word.Text, "AutoHtokey") < 3
        result.Highlight(word)
}

; Credit: iPhilip, Source: https://www.autohotkey.com/boards/viewtopic.php?style=17&p=509167#p509167
; https://en.wikipedia.org/wiki/Levenshtein_distance#Iterative_with_two_matrix_rows
LD(Source, Target, CaseSense := True) {
    if CaseSense ? Source == Target : Source = Target
       return 0
    Source := StrSplit(Source)
    Target := StrSplit(Target)
    if !Source.Length
       return Target.Length
    if !Target.Length
       return Source.Length
    
    v0 := [], v1 := []
    Loop Target.Length + 1
       v0.Push(A_Index - 1)
    v1.Length := v0.Length
    
    for Index, SourceChar in Source {
       v1[1] := Index
       for TargetChar in Target
          v1[A_Index + 1] := Min(v1[A_Index] + 1, v0[A_Index + 1] + 1, v0[A_Index] + (CaseSense ? SourceChar !== TargetChar : SourceChar != TargetChar))
       Loop Target.Length + 1
          v0[A_Index] := v1[A_Index]
    }
    return v1[Target.Length + 1]
 }