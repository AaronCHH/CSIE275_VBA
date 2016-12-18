Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub getArray()

Dim x As Variant: x = Range("D2:D10").Value
MsgBox UBound(x, 1)
MsgBox UBound(x, 2)
Dim i As Integer

For i = LBound(x) To UBound(x)
    Debug.Print x(i, 1)
Next i
End Sub
