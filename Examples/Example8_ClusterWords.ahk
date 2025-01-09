#Requires AutoHotkey v2
#include ..\Lib\OCR.ahk

SetTitleMatchMode 2
result := OCR.FromDesktop()
cluster := OCR.Cluster(result.Words)
for res in cluster
    res.Highlight(-5000)
Sleep 5000