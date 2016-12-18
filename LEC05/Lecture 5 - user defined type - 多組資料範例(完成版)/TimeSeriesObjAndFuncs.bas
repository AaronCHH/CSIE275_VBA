Attribute VB_Name = "TimeSeriesObjAndFuncs"
Option Explicit
Option Base 1

' 為了可以有 Null，我們把 dates 與 values 改為 Variant，但記住:
' dates 裡面要使用日期陣列或 Null
' values 裡面要是 double 陣列或 Null
Public Type TimeSeries
    dates As Variant
    values As Variant
End Type

Public Type MultiTimeSerires
    dates As Variant
    values As Variant
End Type

Public Enum SearchType
    byDate = 0
    byValue = 1
End Enum

Sub createAndPrintTimeSeries()

Dim filePath As Variant: filePath = Application.GetOpenFilename(MultiSelect:=True)
Dim dataFile As Workbook
'避免麻煩，等等讀兩組就好
Dim lastRow As Integer

Set dataFile = Workbooks.Open(filePath(1))
dataFile.Worksheets(1).Activate
lastRow = Cells(2, "A").End(xlDown).Row

Dim data2311 As TimeSeries: data2311 = fromWorksheet(Range("A2", Cells(lastRow, "A")), _
                                                                                                                  Range("B2", Cells(lastRow, "B")))

dataFile.Close

Set dataFile = Workbooks.Open(filePath(2))

dataFile.Worksheets(1).Activate
lastRow = Cells(2, "A").End(xlDown).Row

Dim data2330 As TimeSeries: data2330 = fromWorksheet(Range("A2", Cells(lastRow, "A")), _
                                                                                                                  Range("B2", Cells(lastRow, "B")))

dataFile.Close

Dim mergeData As MultiTimeSerires: mergeData = mergeTimeSeries(data2311, data2330)
Dim printResult As Variant: printResult = multiSeriesToVariantArray(mergeData)
Range("A1", Cells(UBound(printResult), UBound(printResult, 2))).Value = printResult

'Range("A1", Cells(UBound(printResult), "B")).Value = printResult
'printResult = toVariantArray(data2330)
'Range("C1", Cells(UBound(printResult), "D")).Value = printResult
End Sub

' 轉為可以貼製儲存格的Variant陣列
Function toVariantArray(inputData As TimeSeries) As Variant

Dim i As Integer
Dim nData As Integer: nData = UBound(inputData.dates)
ReDim output(nData + 1, 2) As Variant
output(1, 1) = "Date"
output(1, 2) = "Value"

For i = 1 To nData
    output(i + 1, 1) = inputData.dates(i)
    output(i + 1, 2) = inputData.values(i)
Next i

toVariantArray = output
End Function

Function multiSeriesToVariantArray(inputData As MultiTimeSerires) As Variant

Dim i As Integer
Dim j As Integer
Dim nDiffVal As Integer: nDiffVal = UBound(inputData.values, 2)
Dim nData As Integer: nData = UBound(inputData.dates)
ReDim output(nData + 1, nDiffVal + 1) As Variant

output(1, 1) = "Date"
For i = 1 To nDiffVal
    output(1, i + 1) = "Value " & i
Next i

For i = 1 To nData
    output(i + 1, 1) = inputData.dates(i)
    
    For j = 1 To nDiffVal
        output(i + 1, j + 1) = inputData.values(i, j)
    Next j
Next i

multiSeriesToVariantArray = output
End Function


'輸入的函數
Function fromWorksheet(dateRange As Range, valueRange As Range) As TimeSeries

Dim nData As Integer: nData = dateRange.Count

ReDim dates(nData) As Date
ReDim values(nData) As Double
Dim output As TimeSeries

Call rangeToArray(dates, dateRange)
Call rangeToArray(values, valueRange)

output.dates = dates
output.values = values
fromWorksheet = output

End Function

' 輸入陣列
Private Sub rangeToArray(ByRef thisArray As Variant, ByRef inputRange As Range)
Dim inputValues As Variant: inputValues = inputRange.Value
Dim elem As Variant
Dim counts As Integer

For Each elem In inputValues
    counts = counts + 1
    thisArray(counts) = elem
Next elem
End Sub


'選取值落在特定區間內的資料
Function inInterval(inputSeries As TimeSeries, _
                             searchBy As SearchType, _
                             Optional lowerBound As Variant, _
                             Optional upperBound As Variant) As TimeSeries
                              
Dim i As Integer
Dim counts As Integer
Dim thisValue As Variant
Dim nData As Integer: nData = UBound(inputSeries.dates)
ReDim newDates(nData) As Date
ReDim newValues(nData) As Double
                              
Select Case searchBy
    Case byDate
        If IsMissing(lowerBound) Then lowerBound = Application.Min(inputSeries.dates)
        If IsMissing(upperBound) Then upperBound = Application.Max(inputSeries.dates)
        
        For i = 1 To nData
            thisValue = inputSeries.dates(i)
            
            If thisValue >= lowerBound And thisValue <= upperBound Then
                counts = counts + 1
                newDates(counts) = thisValue
                newValues(counts) = inputSeries.values(i)
            End If
        Next i
    Case byValue
        If IsMissing(lowerBound) Then lowerBound = Application.Min(inputSeries.values)
        If IsMissing(upperBound) Then upperBound = Application.Max(inputSeries.values)

        For i = 1 To nData
            thisValue = inputSeries.values(i)
            
            If thisValue >= lowerBound And thisValue <= upperBound Then
                counts = counts + 1
                newDates(counts) = inputSeries.dates(i)
                newValues(counts) = thisValue
            End If
        Next i
End Select

Dim newSeries As TimeSeries

If counts > 0 Then
    ReDim Preserve newDates(counts)
    ReDim Preserve newValues(counts)
    newSeries.dates = newDates
    newSeries.values = newValues
Else
    newSeries.dates = Array(Null)
    newSeries.values = Array(Null)
End If
inInterval = newSeries
End Function

Function mergeTimeSeries(seriesA As TimeSeries, _
                                                      seriesB As TimeSeries) As MultiTimeSerires

Dim valuesA As Variant: valuesA = seriesA.values
Dim valuesB As Variant: valuesB = seriesB.values
Dim datesA As Variant: datesA = seriesA.dates
Dim datesB As Variant: datesB = seriesB.dates
Dim nDataA As Integer: nDataA = UBound(valuesA)
Dim nDataB As Integer: nDataB = UBound(valuesB)
ReDim newDates(nDataA + nDataB) As Date
Dim countsA As Integer: countsA = 1
Dim countsB As Integer: countsB = 1
Dim newCounts As Integer

Do While countsA <= nDataA Or countsB <= nDataB
    newCounts = newCounts + 1
    If datesA(countsA) < datesB(countsB) Then
        newDates(newCounts) = datesA(countsA)
        countsA = countsA + 1
     ElseIf datesA(countsA) > datesB(countsB) Then
        newDates(newCounts) = datesB(countsB)
        countsB = countsB + 1
     Else
        newDates(newCounts) = datesB(countsB)
        countsA = countsA + 1
        countsB = countsB + 1
    End If
Loop
ReDim Preserve newDates(newCounts)
ReDim newValues(newCounts, 2) As Variant

Dim i As Integer
Dim j As Integer
Dim hasData As Boolean

For i = 1 To nDataA
    hasData = False
    For j = 1 To newCounts
        If datesA(i) = newDates(j) Then
            countsA = j
            hasData = True
            Exit For
        End If
    Next j
        
    If hasData Then
        newValues(countsA, 1) = valuesA(i)
    Else
        newValues(countsA, 1) = Null
    End If
Next i
        
For i = 1 To nDataB
    hasData = False
    For j = 1 To newCounts
        If datesB(i) = newDates(j) Then
            countsB = j
            hasData = True
            Exit For
        End If
    Next j
        
    If hasData Then
        newValues(countsB, 2) = valuesB(i)
    Else
        newValues(countsB, 2) = Null
    End If
Next i

Dim output As MultiTimeSerires
output.dates = newDates
output.values = newValues
mergeTimeSeries = output
End Function























