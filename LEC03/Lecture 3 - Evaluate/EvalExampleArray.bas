Attribute VB_Name = "EvalExampleArray"
Option Explicit
Option Base 1

Sub createArray()

On Error GoTo isOneDIm
Dim x As Variant: x = Array(1, 3, 5, 7, 9) '二維陣列怎麼辦 ?

MsgBox "第一維長度: " & UBound(x, 1)
MsgBox "第二維長度: " & UBound(x, 2)
Exit Sub

isOneDIm:
MsgBox "非二維陣列"

End Sub
