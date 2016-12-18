Attribute VB_Name = "HistoricalVolatilityExample"
Public Sub calculatLogPrice()

Dim i As Integer, logPrice() As Double
ReDim logPrice(0 To 60)


For i = 0 To 60
   logPrice(i) = Log(Cells(i + 2, 7).Value)
   Range("L" & (3 + i)).Value = logPrice(i)
Next i

End Sub

Public Sub calculateLogRetrun()

Dim i As Integer, logReturn() As Double, deltaT As Double
ReDim logReturn(1 To 60)

If Range("L3").Value = "" Then
    MsgBox ("請先完成Step 1.")
    Exit Sub
End If

deltaT = 1 / 365

For i = 1 To 60
   logReturn(i) = (Range("L" & (2 + i)).Value - Range("L" & (3 + i)).Value) / deltaT ^ 0.5
   
   Range("O" & (2 + i)).Value = logReturn(i)
Next i



End Sub

Public Sub calculateAverageLogReturn()

Dim i As Double, sumLogReturn As Double, averageLogReturn As Double


If Range("O3").Value = "" Then
    MsgBox ("請先完成Step 2.")
    Exit Sub
End If

sumLogReturn = 0

For i = 1 To 60
  sumLogReturn = sumLogReturn + Range("O" & (2 + i)).Value
Next i

averageLogReturn = sumLogReturn / 60

Range("R3").Value = averageLogReturn

End Sub

Public Sub calculateSquareVariation()

Dim i As Integer, squareVariation() As Double, averageLogReturn As Double
ReDim squareVariation(1 To 60)

If Range("R3").Value = "" Then
    MsgBox ("請先完成Step 3.")
    Exit Sub
End If

averageLogReturn = Range("R3").Value

For i = 1 To 60
   squareVariation(i) = (Range("O" & (2 + i)).Value - averageLogReturn) ^ 2
   Range("U" & (2 + i)).Value = squareVariation(i)
Next i



End Sub

Public Sub calculateHistoricalVolatility()

Dim sumVariation As Double, volatility As Double

If Range("U3").Value = "" Then
    MsgBox ("請先完成Step 4.")
    Exit Sub
End If

For i = 1 To 60
    sumVariation = sumVariation + Range("U" & (2 + i)).Value
Next i

volatility = (sumVariation / (60 - 1)) ^ 0.5

Range("X3").Value = volatility

End Sub

Public Sub reset()
    Range("L3:L63").ClearContents
    Range("O3:O62").ClearContents
    Range("U3:U62").ClearContents
    Range("R3").ClearContents
    Range("X3").ClearContents
End Sub

