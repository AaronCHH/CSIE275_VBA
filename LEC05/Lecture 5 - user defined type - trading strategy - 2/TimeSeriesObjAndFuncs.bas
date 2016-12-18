Attribute VB_Name = "TimeSeriesObjAndFuncs"
Option Explicit
Option Base 1

' ���F�i�H�� Null�A�ڭ̧� dates �P values �אּ Variant�A���O��:
' dates �̭��n�ϥΤ���}�C�� Null
' values �̭��n�O double �}�C�� Null
Public Type TimeSeries
    dates As Variant
    values As Variant
    varName As String
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
                                                                                                                                                              Range("B2", Cells(lastRow, "B")), _
                                                                                                                                                              "2311 adj close")
dataFile.Close

Dim returnSeries As TimeSeries
returnSeries = calculateReturnRate(adjColsedTimeSeries)
Dim arParameters() As Double: arParameters = setARModel(returnSeries, 1)

For i = 1 To UBound(arParameters)
    Debug.Print arParameters(i)
Next i

'�C�L���G
Dim printResult As Variant: printResult = toVariantArray(returnSeries)
Range("A1", Cells(UBound(printResult), "B")).Value = printResult
End Sub

' �ର�i�H�K�s�x�s�檺Variant�}�C
Function toVariantArray(inputData As TimeSeries) As Variant

Dim i As Integer
Dim nData As Integer: nData = UBound(inputData.dates)
ReDim output(nData + 1, 2) As Variant
output(1, 1) = "Date"
output(1, 2) = inputData.varName
For i = 1 To nData
    output(i + 1, 1) = inputData.dates(i)
    output(i + 1, 2) = inputData.values(i)
Next i

toVariantArray = output
End Function

'��J�����
Function fromWorksheet(dateRange As Range, _
                                                   valueRange As Range, _
                                                   Optional varName As String = "Value") As TimeSeries

Dim nData As Integer: nData = dateRange.Count

ReDim dates(nData) As Date
ReDim values(nData) As Double
Dim output As TimeSeries

Call rangeToArray(dates, dateRange)
Call rangeToArray(values, valueRange)

output.dates = dates
output.values = values
output.varName = varName
fromWorksheet = output

End Function

' ��J�}�C
Private Sub rangeToArray(ByRef thisArray As Variant, ByRef inputRange As Range)
Dim inputValues As Variant: inputValues = inputRange.Value
Dim elem As Variant
Dim counts As Integer

For Each elem In inputValues
    counts = counts + 1
    thisArray(counts) = elem
Next elem
End Sub


'����ȸ��b�S�w�϶��������
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
newSeries.varName = inputSeries.varName
inInterval = newSeries
End Function

Function calculateReturnRate(inputSeries As TimeSeries, _
                                                           Optional nPeriod As Integer = 1) As TimeSeries

Dim nData As Integer: nData = UBound(inputSeries.dates)
ReDim returnDates(nData - nPeriod) As Date
ReDim returnRates(nData - nPeriod) As Double
Dim i As Integer

For i = (1 + nPeriod) To nData
    returnDates(i - nPeriod) = inputSeries.dates(i)
    returnRates(i - nPeriod) = Log(inputSeries.values(i)) - Log(inputSeries.values(i - nPeriod))
Next i

Dim returnSeries As TimeSeries
returnSeries.dates = returnDates
returnSeries.values = returnRates
returnSeries.varName = inputSeries.varName + " Retrun Rate"
calculateReturnRate = returnSeries
End Function


Function setARModel(inputSeries As TimeSeries, order As Integer) As Double()

Dim dataValues As Variant: dataValues = inputSeries.values
Dim nOriginalData As Integer: nOriginalData = UBound(inputSeries.dates)
Dim nNewData As Integer: nNewData = nOriginalData - order
Dim i As Integer
Dim j As Integer

ReDim y(nNewData) As Double
ReDim x(order, nNewData) As Double

For i = 1 To nNewData
    y(i) = dataValues(i + order)
    
    For j = 1 To order
        x(j, i) = dataValues(i + order - j)
    Next j
Next i

Dim regressionResult As Variant: regressionResult = Application.LinEst(y, x, True, False)
Dim beta As Variant
Dim nParam As Integer: nParam = order + 1
ReDim betaVec(nParam) As Double

For Each beta In regressionResult
    betaVec(nParam) = beta
    nParam = nParam - 1
Next beta

setARModel = betaVec
End Function

















