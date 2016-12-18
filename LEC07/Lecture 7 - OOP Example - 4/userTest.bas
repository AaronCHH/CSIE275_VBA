Attribute VB_Name = "userTest"
Option Explicit
Option Base 1


Sub test()
Dim spotPrice As TimeSeries
ReDim dates(3) As Date: dates(1) = DateValue("2013/12/13")
ReDim values(3) As Double: values(1) = 102
Dim i As Integer

For i = 2 To 3
    dates(i) = dates(i - 1) + 1
    values(i) = values(i - 1) + Rnd * 1.5
Next i

spotPrice = fromArray(dates, values, "spot")

Dim instrument As PlainVanilla
Set instrument = newPlainVanilla(100, DateValue("2013/12/14"), callOption)

Dim payoffSeries As TimeSeries
payoffSeries = instrument.calculatePayoff(spotPrice)

Range("A1:B4").value = toVariantArray(payoffSeries)

End Sub







