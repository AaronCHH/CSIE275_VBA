Attribute VB_Name = "EvalExamplePmax"
Option Explicit
Option Base 1

Function pmax(x As Range, y As Range) As Double()

Dim nRow As Integer: nRow = x.Rows.Count
Dim nCol As Integer: nCol = x.Columns.Count
Dim xValue As Variant: xValue = x.Value
Dim yValue As Variant: yValue = y.Value
Dim i As Integer
Dim j As Integer
ReDim output(nRow, nCol) As Double

For i = 1 To nRow
    For j = 1 To nCol
        If xValue(i, j) >= yValue(i, j) Then 
          output(i, j) = xValue(i, j) 
        Else output(i, j) = yValue(i, j)
    Next j
Next i
      
pmax = output
End Function












