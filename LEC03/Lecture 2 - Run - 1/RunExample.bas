Attribute VB_Name = "RunExample"
Option Explicit
Option Base 1

Function simpleMovingAverage(stockPrices As Variant) As Variant

simpleMovingAverage = Application.Average(stockPrices)

End Function

Function weightedMovingAverage(stockPrices As Variant) As Variant

Dim dominator As Integer
Dim counts As Integer
Dim elem As Variant

For Each elem In stockPrices
    counts = counts + 1
    dominator = dominator + counts
    weightedMovingAverage = weightedMovingAverage + elem * counts
Next elem

weightedMovingAverage = weightedMovingAverage / dominator

End Function


Function movingAverage(stockPrices As Variant, method As String) As Variant

'Select Case method
'    Case "simpleMovingAverage"
'        movingAverage = simpleMovingAverage(stockPrices)
'    Case "weightedMovingAverage"
'        movingAverage = weightedMovingAverage(stockPrices)
'End Select
movingAverage = Application.Run(method, stockPrices)

End Function






















