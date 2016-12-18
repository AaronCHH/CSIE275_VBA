Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
'Version 2
Dim alpha As Variant

Sub setAlpha()
alpha = Range("Alpha").Value
MsgBox "比率重新設定"
End Sub

Function emaCalculator(priceArray As Range) As Variant
Dim dominator As Variant                                         '加權平均加總所除上的分母
Dim onePrice As Range                                            '某天價格的儲存格
Dim weight As Variant                                                                       '該價格之權重
Dim counts As Integer: counts = priceArray.Count                 '現在權重的 alpha 之次方

For Each onePrice In priceArray
    counts = counts - 1
    weight = alpha ^ counts
    emaCalculator = emaCalculator + onePrice.Value * weight
    dominator = dominator + weight
Next onePrice

emaCalculator = emaCalculator / dominator
End Function
