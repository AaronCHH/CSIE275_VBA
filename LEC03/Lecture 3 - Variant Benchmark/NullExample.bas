Attribute VB_Name = "NullExample"
Option Explicit
Option Base 1

Sub nullExample()

Dim x As Variant: x = Null
Dim y As Variant: y = 1
Dim z As Variant

z = Nz(x > y)
y = x + y

MsgBox z
MsgBox IsNull(z)
End Sub
