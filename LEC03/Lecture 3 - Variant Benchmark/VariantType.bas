Attribute VB_Name = "VariantType"
Option Explicit
Option Base 1

Sub getVariableType()

Dim x As Variant

'¿é¤J x = 100
x = 100
MsgBox "x: " & x & Chr(10) & "Type of x: " & VarType(x)

'¿é¤J x =500000
x = 500000
MsgBox "x: " & x & Chr(10) & "Type of x: " & VarType(x)

'¿é¤J x = 1E+5
x = 100000#
MsgBox "x: " & x & Chr(10) & "Type of x: " & VarType(x)


End Sub
