Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub goToExample()

Dim lastRow As Integer: lastRow = Cells(1, 1).End(xlDown).Row
Dim rateValues As Variant: rateValues = Range("B2", Cells(lastRow, "B")).Value
Dim counts As Integer: counts = 0
Dim firstTenGreaterThanSix(10) As Double
Dim i  As Integer

For i = 2 To lastRow
    If rateValues(i - 1, 1) > 6 Then
        counts = counts + 1
        firstTenGreaterThanSix(counts) = rateValues(i - 1, 1)
    End If
    
    If counts = 10 Then GoTo getTenElements
Next i
'�S���Q�Ӫ���:
MsgBox "�`���� " & counts & " �����"
Exit Sub
'�T��`���Q�Ӫ�����:
getTenElements:
    MsgBox "�T��`����Q��"
    

End Sub

