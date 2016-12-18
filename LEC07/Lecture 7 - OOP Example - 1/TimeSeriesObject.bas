Attribute VB_Name = "TimeSeriesObject"
Option Explicit
Option Base 1

Public Type TimeSeries
    dates As Variant
    values As Variant
    varName As String
End Type

Public Enum SearchType
    byDate = 0
    byValue = 1
End Enum
' 轉為可以貼製儲存格的Variant陣列
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

'輸入的函數
Function fromArray(dateArray As Variant, _
                               valueArray As Variant, _
                               Optional varName As String = "Value") As TimeSeries

    Dim output As TimeSeries
    output.dates = dateArray
    output.values = valueArray
    output.varName = varName
    
    If isConsistent(output) Then
        fromArray = output
    Else
        Err.Raise Number:=11001, Source:="fromArray", Description:="日期與值長度不一"
    End If
    
End Function

' 輸入陣列
Private Sub rangeToArray(ByRef thisArray As Variant, ByRef inputRange As Range)
    Dim inputValues As Variant: inputValues = inputRange.value
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
            If IsMissing(lowerBound) Then
                lowerBound = inputSeries.dates(LBound(inputSeries.dates))
            Else
                lowerBound = DateValue(lowerBound)
            End If
            
            If IsMissing(upperBound) Then
                upperBound = inputSeries.dates(UBound(inputSeries.dates))
            Else
                upperBound = DateValue(upperBound)
            End If
            
            For i = 1 To nData
                thisValue = inputSeries.dates(i)
                If (thisValue >= lowerBound) And (thisValue <= upperBound) Then
                    counts = counts + 1
                    newDates(counts) = thisValue
                    newValues(counts) = inputSeries.values(i)
                End If
            Next i
        Case byValue
            If IsMissing(lowerBound) Then lowerBound = inputSeries.values(LBound(inputSeries.values))
            If IsMissing(upperBound) Then upperBound = inputSeries.values(UBound(inputSeries.values))
    
            For i = 1 To nData
                thisValue = inputSeries.values(i)
                
                If (thisValue >= lowerBound) And (thisValue <= upperBound) Then
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

Public Function isConsistent(inputSeries As TimeSeries) As Boolean

    isConsistent = True
    
On Error GoTo notConsistent:

    If UBound(inputSeries.values) <> UBound(inputSeries.dates) Then
        isConsistent = False
    End If
    
    Exit Function
 
notConsistent
    isConsistent = False
    
End Function

























