Attribute VB_Name = "RunExample"
Option Explicit
Option Base 1

Function simpleMovingAverage(stockPrices As Variant) As Variant

simpleMovingAverage = appliction.Average(stockPrices)

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

movingAverage = Application.Run(method, stockPrices)

End Function






















