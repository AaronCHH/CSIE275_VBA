Attribute VB_Name = "EvalExampleArray"
Option Explicit
Option Base 1

Sub createArray()

On Error GoTo isOneDIm
Dim x As Variant: x = Array(1, 3, 5, 7, 9) '�G���}�C���� ?

MsgBox "�Ĥ@������: " & UBound(x, 1)
MsgBox "�ĤG������: " & UBound(x, 2)
Exit Sub

isOneDIm:
MsgBox "�D�G���}�C"

End Sub
