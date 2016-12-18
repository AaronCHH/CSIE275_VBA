Attribute VB_Name = "Module1"
Option Explicit

Function mySum(inputData As Variant) As Double

Dim oneVal As Variant

mySum = 0

For Each oneVal In inputData
    mySum = mySum + oneVal
Next oneVal

End Function


'ParamArray Version
Function mySum_2(ParamArray inputData()) As Variant

Dim oneElem As Variant, _
    elemInElem As Variant

mySum_2 = 0

For Each oneElem In inputData

    '不是陣列就直接累加
    If Not IsArray(oneElem) Then
        mySum_2 = mySum_2 + oneElem
        
    '是陣列就再用一層For Each 取值累加
    Else
        For Each elemInElem In oneElem
            mySum_2 = mySum_2 + elemInElem
        Next elemInElem
        
    End If
    
Next oneElem

End Function