Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
Option Base 1
'Version 2
'�� alpha �ܥ����ܼ�
Public alpha As Variant

Function getAlpha()
getAlpha = alpha
End Function

Sub setAlpha()
alpha = Range("Alpha").Value
MsgBox "��v���s�]�w"
End Sub

Function emaCalculator(priceArray As Range, _
                                                Optional showAlpha As Boolean = True) As Variant
Dim dominator As Variant                                                                '�[�v�����[�`�Ұ��W������
Dim onePrice As Range                                                                     '�Y�ѻ��檺�x�s��
Dim weight As Variant                                                                       '�ӻ��椧�v��
Dim counts As Integer: counts = priceArray.Count                 '�{�b�v���� alpha ������
Dim nOutput As Integer

If showAlpha Then
    ReDim output(2) As Double
    output(2) = alpha
Else
    ReDim output(1) As Double
End If

For Each onePrice In priceArray
    counts = counts - 1
    weight = alpha ^ counts
    output(1) = output(1) + onePrice.Value * weight
    dominator = dominator + weight
Next onePrice

output(1) = output(1) / dominator

emaCalculator = output
End Function










