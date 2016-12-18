Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
Option Base 1
'Version 2
'р alpha 跑办跑计
Public alpha As Variant

Function getAlpha()
getAlpha = alpha
End Function

Sub setAlpha()
alpha = Range("Alpha").Value
MsgBox "ゑ瞯穝砞﹚"
End Sub

Function emaCalculator(priceArray As Range, _
                                                Optional showAlpha As Boolean = True) As Variant
Dim dominator As Variant                                                                '舦キА羆┮埃だダ
Dim onePrice As Range                                                                     '琘ぱ基纗
Dim weight As Variant                                                                       '赣基ぇ舦
Dim counts As Integer: counts = priceArray.Count                 '瞷舦 alpha ぇΩよ
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










