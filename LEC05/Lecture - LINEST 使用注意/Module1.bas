Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub linestExample()

Dim y As Variant: y = Range("A2:A10").Value
Dim x As Variant: x = Range("B2:C10").Value
Dim i As Integer

Dim betaVec As Variant: betaVec = Application.LinEst(y, x, True, False)
Dim beta As Variant
For Each beta In betaVec
    Debug.Print beta
Next beta
End Sub
