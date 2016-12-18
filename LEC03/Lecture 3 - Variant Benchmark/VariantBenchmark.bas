Attribute VB_Name = "VariantBenchmark"
Option Explicit
Option Base 1
Dim sampleDataString() As String
Dim sampleDataVariant As Variant


Private Function calculateGoodRatioViaVariant(samples)

Dim nSample: nSample = UBound(samples)
Dim i

For i = 1 To nSample
    Select Case samples(i)
        Case "GOOD"
            calculateGoodRatioViaVariant = calculateGoodRatioViaVariant + 1
    End Select
Next i
    
calculateGoodRatioViaVariant = calculateGoodRatioViaVariant / nSample

End Function

Private Function calculateGoodRatioViaString(samples() As String) As Double

Dim nSample As Long: nSample = UBound(samples)
Dim i As Long

For i = 1 To nSample
    Select Case samples(i)
        Case "GOOD"
            calculateGoodRatioViaString = calculateGoodRatioViaString + 1
    End Select
Next i
    
calculateGoodRatioViaString = calculateGoodRatioViaString / nSample

End Function

Private Sub getData()

Dim nData As Long: nData = Cells(1, 1).End(xlDown).Row - 1
Dim dataInWorksheets As Variant: dataInWorksheets = Range(Cells(2, 2), Cells(nData + 1, 2)).Value
Dim i As Long

ReDim sampleDataString(nData)

For i = 1 To nData
    sampleDataString(i) = dataInWorksheets(i, 1)
Next i

sampleDataVariant = sampleDataString

End Sub
