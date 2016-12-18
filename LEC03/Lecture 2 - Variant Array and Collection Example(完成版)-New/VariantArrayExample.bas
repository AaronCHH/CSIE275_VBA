Attribute VB_Name = "VariantArrayExample"
Option Explicit
Option Base 1

Sub createPortfolioArray()
Application.ScreenUpdating = False

Dim historicalDataSheet As Worksheet
Set historicalDataSheet = Worksheets("���v��ƭ���")

historicalDataSheet.Activate

Dim lastRow As Integer
lastRow = Cells(1, 1).End(xlDown).Row

Dim nStock As Integer
nStock = Cells(1, 1).End(xlToRight).Column - 1

Dim stockNames As Variant
stockNames = Range("B1", Cells(1, nStock + 1)).Value

Dim presentValueDate As Date
presentValueDate = Worksheets("���զX�{��").Range("B1").Value

Dim dateVec As Variant
dateVec = Range("A2", Cells(lastRow, "A")).Value

Dim counts As Integer
Dim thisName As Variant
ReDim portfolioArray(nStock)

For Each thisName In stockNames
    counts = counts + 1
    portfolioArray(counts) = dataToArray(thisName, dateVec, presentValueDate)
Next thisName

printHistoricalData (portfolioArray(3))

End Sub

Function dataToArray(oneStockName As Variant, _
                                dateVec As Variant, _
                                presentValueDate As Date) As Variant()
Dim i As Integer
Dim cellCol As Integer
Dim findCell
'�}�C���׬��|�A��J�|�ո��
Dim stockData(4) As Variant

'���i�J�{�Ȥu�@��A��J���
Worksheets("���զX�{��").Activate
Set findCell = Cells.Find(oneStockName)
cellCol = findCell.Column
stockData(1) = presentValueDate
stockData(2) = Cells(5, cellCol).Value '�ĤG�Ӥ�����J�{��


'���o���v��ƪ���
Dim lastRow As Integer: lastRow = Worksheets("���v��ƭ���"). _
                                                                Cells(1, 1).End(xlDown).Row

'�i�J���v��Ƥu�@��A���o�ɶ��ǦC
Worksheets("���v��ƭ���").Activate
cellCol = Cells.Find(oneStockName).Column
stockData(3) = dateVec
stockData(4) = Range(Cells(2, cellCol), Cells(lastRow, cellCol)).Value '�ĥ|�Ӥ�����J���v�ѻ�

dataToArray = stockData

End Function

'�g�@����
Sub printHistoricalData(inputData As Variant)
Dim nData As Integer: nData = UBound(inputData(3))
Dim i As Integer

Debug.Print "���     " + "      ���v����"

For i = 1 To nData
    Debug.Print inputData(3)(i, 1) & "    " & inputData(4)(i, 1)
Next i

End Sub















