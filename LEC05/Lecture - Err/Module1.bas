Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Dim collectionSheet As Worksheet
Dim dataFile As Workbook
Dim dataPathCollection As Variant

Sub resumeExample()

Set collectionSheet = Worksheets("Example")
dataPathCollection = Application.GetOpenFilename(MultiSelect:=True)

Dim oneFilePath As Variant
Dim colCounts As Integer

On Error Resume Next ' 錯了就換下一個

For Each oneFilePath In dataPathCollection
    colCounts = colCounts + 1
    Call inputNextData(oneFilePath, colCounts)
Next oneFilePath

End Sub

Private Sub inputNextData(filePath As Variant, inputCol As Integer)
' 在此處加入 On Error GoTo 吧!!
Dim lastRow As Integer

collectionSheet.Cells(1, inputCol).Value = "Data " & inputCol
Set dataFile = Workbooks.Open(filePath)
lastRow = dataFile.Worksheets(1).Cells(1, 1).End(xlDown).Row
dataFile.Worksheets(1).Range("A2", Cells(lastRow, "A")).Copy Destination:=collectionSheet.Cells(2, inputCol)
dataFile.Close
End Sub












