Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
Option Base 1
'Version 3
'靜態變數版


Private Sub setAlpha(ByRef alpha)
alpha = Range("Alpha").Value
MsgBox "比率重新設定"
End Sub


Function emaCalculator(ParamArray priceArray()) As Variant
Dim dominator As Variant                           '加權平均加總所除上的分母
Dim oneElem As Variant                             '陣列的某個元素
Dim elemInElem As Variant                          ' 元素中的元素
Dim weight As Variant                                                                       '該價格之權重
Dim counts As Variant                                                                       '現在權重的 alpha 之次方

' ParamArray 版本要計算數量會比較麻煩，需要跑迴圈

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




