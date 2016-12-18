Attribute VB_Name = "InterpolationMethods"
Option Explicit
Option Base 1

' plot rates for 1D to 1Y by dates
Sub plotResult()

Worksheets("Interpolation").Activate

Dim nData As Integer: nData = Cells(1, 1).End(xlDown).Row - 1
Dim i As Integer
Dim yearBase As Double: yearBase = 1 / 360
Dim yearFraction As Double
Dim chartName As String

ReDim rates(nData) As Double, _
      maturities(nData) As Double, _
      xHat(365) As Double, _
      yHat(365) As Double

For i = 1 To nData
    rates(i) = Cells(i + 1, "C").Value
    maturities(i) = Cells(i + 1, "B").Value
Next i

i = 0
For yearFraction = yearBase To 365 * yearBase Step yearBase
    i = i + 1
    xHat(i) = yearFraction
    yHat(i) = interpolation(xHat(i), rates, maturities, "linerInterpolation")
Next yearFraction

Charts.Add before:=Worksheets("Interpolation")
chartName = "Interpolation Result " & Charts.Count

With ActiveChart
     .Name = chartName
     .SeriesCollection.NewSeries
     .SeriesCollection(1).Values = yHat
     .SeriesCollection(1).XValues = xHat
     .SeriesCollection(1).ChartType = xlLine
End With
     

End Sub

Private Function interpolation(targetMaturity As Double, _
                                                        rates() As Double, _
                                                        maturities() As Double, _
                                                        method As String) As Double
                                                       
Dim nData As Integer: nData = UBound(rates)

If targetMaturity <= maturities(1) Then

    interpolation = rates(1)
    
ElseIf targetMaturity >= maturities(nData) Then

    interpolation = rates(nData)
    
Else
    Dim nextIndex As Integer
    Dim singleMaturity As Variant
    
    For Each singleMaturity In maturities
    
        nextIndex = nextIndex + 1
        
        If singleMaturity > targetMaturity Then
            Exit For
        End If
    Next singleMaturity
    
    Dim preRate As Double: preRate = rates(nextIndex - 1)
    Dim nextRate As Double: nextRate = rates(nextIndex)
    Dim preMat As Double: preMat = maturities(nextIndex - 1)
    Dim nextMat As Double: nextMat = maturities(nextIndex)
   
   interpolation = Application.Run(method, targetMaturity, preRate, _
                                                             nextRate, preMat, nextMat)
End If
                                                       
End Function

' 內插單一利率之作法，rates 與 maturities 兩者都為一維陣列
Private Function nearestNeighbor(targetMaturity As Double, _
                                                              preRate As Double, _
                                                              nextRate As Double, _
                                                              preMat As Double, _
                                                              nextMat As Double) As Double
                                                              
 If Abs(targetMaturity - preMat) <= Abs(targetMaturity - nextMat) Then
        nearestNeighbor = preRate
    Else
        nearestNeighbor = nextRate
End If
End Function

' 內插單一利率之作法，rates 與 maturities 兩者都為一維陣列
Private Function linerInterpolation(targetMaturity As Double, _
                                                              preRate As Double, _
                                                              nextRate As Double, _
                                                              preMat As Double, _
                                                              nextMat As Double) As Double
                                                              
 linerInterpolation = preRate + (nextRate - preRate) * _
                                        (targetMaturity - preMat) / _
                                        (nextMat - preMat)
End Function



