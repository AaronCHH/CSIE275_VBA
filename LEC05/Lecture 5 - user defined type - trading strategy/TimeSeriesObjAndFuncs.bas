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

Public Enum SearchType
    byDate = 0
    byValue = 1
End Enum

Sub createAndPrintTimeSeries()

Dim filePath As String: filePath = Application.GetOpenFilename()
Dim dataFile As Workbook: Set dataFile = Workbooks.Open(filePath)
Dim i As Integer
Dim lastRow As Integer

dataFile.Worksheets(1).Activate
lastRow = Cells(2, "A").End(xlDown).Row

Dim adjColsedTimeSeries As TimeSeries: adjColsedTimeSeries = fromWorksheet(Range("A2", Cells(lastRow, "A")), _
                                                                                                                        Range("B2", Cells(lastRow, "B")))

dataFile.Close

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

