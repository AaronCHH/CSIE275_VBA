Attribute VB_Name = "MontCarloMethodExample"
Sub generateRandomNumber()

Dim i As Integer, j As Integer, n As Integer

For n = 1 To 4
    If Range("B" & (3 + n)).Value = "" Then
        MsgBox ("請先完成Step 1.")
        Exit Sub
    End If
Next n
    

For i = 0 To 2
    For j = 1 To 4
        Cells(14 + i, 2 + j).Value = Rnd()
    Next j
Next i
    
End Sub
Sub standardNormalInverse()

Dim i As Integer, j As Integer, m As Integer, n As Integer



For m = 0 To 2
    For n = 1 To 4
        If Cells(14 + m, 2 + n).Value = "" Then
            MsgBox ("請先完成Step 2.")
            Exit Sub
        End If
    Next n
Next m
    


For i = 0 To 2
    For j = 1 To 4
        Cells(24 + i, 2 + j).Value = Application.WorksheetFunction.NormSInv(Cells(14 + i, 2 + j).Value)
    Next j
Next i

End Sub

Sub calculateMonthlyReturn()

Dim i As Integer, j As Integer, randomWalk As Double, deltaT As Double, riskFreeRate As Double, volatility As Double, m As Integer, n As Integer

For m = 0 To 2
    For n = 1 To 4
        If Cells(24 + m, 2 + n).Value = "" Then
            MsgBox ("請先完成Step 3.")
            Exit Sub
        End If
    Next n
Next m

riskFreeRate = Range("B4").Value
volatility = Range("B6").Value
deltaT = 1 / 12

Range("B33:B35").Value = 1

For i = 0 To 2
    For j = 1 To 4
        randomWalk = Cells(24 + i, 2 + j).Value
        Cells(33 + i, 2 + j).Value = Exp((riskFreeRate - volatility * volatility / 2) * deltaT + _
                                        volatility * randomWalk * deltaT ^ 0.5)
    Next j
Next i

End Sub

Sub SimulationExample()

Dim i As Integer, j As Integer, initialAssetprice As Double, m As Integer, n As Integer


For m = 0 To 2
    For n = 1 To 4
        If Cells(33 + m, 2 + n).Value = "" Then
            MsgBox ("請先完成Step 4.")
            Exit Sub
        End If
    Next n
Next m

initialAssetprice = Range("B7")
     

i = 0
j = 0

If Range("F44") <> "" Then
    MsgBox ("已全部模擬完畢")
      Exit Sub
End If

Do While Cells((42 + i), 6) <> ""
    If i < 2 Then
        i = i + 1
    End If
Loop

Do While Cells((42 + i), (2 + j)) <> ""
    If i < 4 Then
        j = j + 1
    End If
Loop
    
If j = 0 Then
    Cells((42 + i), (2 + j)).Value = initialAssetprice
    MsgBox ("模擬第 " & (i + 1) & " 條路徑開始")
Else
    Cells((42 + i), (2 + j)).Value = Cells((42 + i), (1 + j)).Value * Cells(33 + i, 2 + j).Value
    MsgBox ("模擬第 " & (i + 1) & " 條路徑，第 " & j & " 個月價格")
End If



End Sub

Sub simulationAgain()
    Range("B14:F16").ClearContents
    Range("B24:F26").ClearContents
    Range("B33:F35").ClearContents
    Range("B42:F44").ClearContents
End Sub
