Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
'Version 2
Dim alpha As Variant

Sub setAlpha()
alpha = Range("Alpha").Value
MsgBox "��v���s�]�w"
End Sub

Function emaCalculator(priceArray As Range) As Variant
Dim dominator As Variant                                         '�[�v�����[�`�Ұ��W������
Dim onePrice As Range                                            '�Y�ѻ��檺�x�s��
Dim weight As Variant                                                                       '�ӻ��椧�v��
Dim counts As Integer: counts = priceArray.Count                 '�{�b�v���� alpha ������

For Each onePrice In priceArray
    counts = counts - 1
    weight = alpha ^ counts
    emaCalculator = emaCalculator + onePrice.Value * weight
    dominator = dominator + weight
Next onePrice

emaCalculator = emaCalculator / dominator
End Function
