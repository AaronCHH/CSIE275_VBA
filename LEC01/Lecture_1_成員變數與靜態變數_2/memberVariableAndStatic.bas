Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
Option Base 1
'Version 2
Dim alpha As Variant

Function getAlpha()
getAlpha = alpha
End Function

Sub setAlpha()
alpha = Range("Alpha").Value
MsgBox "比率重新設定"
End Sub

Function emaCalculator(priceArray As Range, _
                                                Optional showAlpha As Boolean = True) As Variant
Dim dominator As Variant                                                                '加權平均加總所除上的分母
Dim onePrice As Range                                                                     '某天價格的儲存格
Dim weight As Variant                                                                       '該價格之權重
Dim counts As Integer: counts = priceArray.Count                 '現在權重的 alpha 之次方
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










