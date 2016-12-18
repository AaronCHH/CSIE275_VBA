Attribute VB_Name = "VariantArrayExample"
Option Explicit
Option Base 1

Sub createPortfolioArray()
Application.ScreenUpdating = False

Dim historicalDataSheet As Worksheet
Set historicalDataSheet = Worksheets("歷史資料頁面")

historicalDataSheet.Activate

Dim lastRow As Integer
lastRow = Cells(1, 1).End(xlDown).Row

Dim nStock As Integer
nStock = Cells(1, 1).End(xlToRight).Column - 1

Dim stockNames As Variant
stockNames = Range("B1", Cells(1, nStock + 1)).Value

Dim presentValueDate As Date
presentValueDate = Worksheets("投資組合現值").Range("B1").Value

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
'陣列長度為四，放入四組資料
Dim stockData(4) As Variant

'先進入現值工作表，輸入資料
Worksheets("投資組合現值").Activate
Set findCell = Cells.Find(oneStockName)
cellCol = findCell.Column
stockData(1) = presentValueDate
stockData(2) = Cells(5, cellCol).Value '第二個元素放入現值


'取得歷史資料長度
Dim lastRow As Integer: lastRow = Worksheets("歷史資料頁面"). _
                                                                Cells(1, 1).End(xlDown).Row

'進入歷史資料工作表，取得時間序列
Worksheets("歷史資料頁面").Activate
cellCol = Cells.Find(oneStockName).Column
stockData(3) = dateVec
stockData(4) = Range(Cells(2, cellCol), Cells(lastRow, cellCol)).Value '第四個元素放入歷史股價

dataToArray = stockData

End Function

'寫一支函數
Sub printHistoricalData(inputData As Variant)
Dim nData As Integer: nData = UBound(inputData(3))
Dim i As Integer

Debug.Print "日期     " + "      歷史價格"

For i = 1 To nData
    Debug.Print inputData(3)(i, 1) & "    " & inputData(4)(i, 1)
Next i

End Sub















