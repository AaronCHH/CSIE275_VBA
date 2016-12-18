Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
Option Base 1
'Version 3
'�R�A�ܼƪ�


Private Sub setAlpha(ByRef alpha)
alpha = Range("Alpha").Value
MsgBox "��v���s�]�w"
End Sub


Function emaCalculator(ParamArray priceArray()) As Variant
Dim dominator As Variant                           '�[�v�����[�`�Ұ��W������
Dim oneElem As Variant                             '�}�C���Y�Ӥ���
Dim elemInElem As Variant                          ' ������������
Dim weight As Variant                                                                       '�ӻ��椧�v��
Dim counts As Variant                                                                       '�{�b�v���� alpha ������

' ParamArray �����n�p��ƶq�|����·СA�ݭn�]�j��

For Each oneElem In priceArray
    If IsArray(oneElem) Then
        counts = counts + UBound(Application.Transpose(oneElem))
    Else
        counts = counts + 1
    End If
Next oneElem

Static isInitialized As Boolean
Static alpha As Variant

If Not isInitialized Then
    setAlpha (alpha)
    isInitialized = True
End If

For Each oneElem In priceArray
    If Not IsArray(oneElem) Then
        counts = counts - 1
        weight = alpha ^ counts
        emaCalculator = emaCalculator + oneElem * weight
        dominator = dominator + weight
    Else
        For Each elemInElem In oneElem
            counts = counts - 1
            weight = alpha ^ counts
            emaCalculator = emaCalculator + elemInElem * weight
            dominator = dominator + weight
        Next elemInElem
    End If
Next oneElem

emaCalculator = emaCalculator / dominator
End Function




