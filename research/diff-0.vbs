Option Explicit

On Error Resume Next

Diff

Sub Diff

Dim objWord
Dim objDoc
Dim f0
Dim f1

f0 = "D:\\Docs\\1.doc"
f1 = "D:\\Docs\\2.doc"

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

objWord.Documents.Open(f0)

objWord.ActiveDocument.Merge f1

' , MergeTarget:=wdMergeTargetCurrent, DetectFormatChanges:=True

objWord.ActiveDocument.Close

End Sub
