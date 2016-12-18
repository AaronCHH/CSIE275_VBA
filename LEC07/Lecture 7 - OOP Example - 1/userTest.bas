Attribute VB_Name = "userTest"
Option Explicit
Option Base 1


Sub test()

Dim instrument1 As New PlainVanilla
instrument1.strike = 100
instrument1.expiryDate = DateValue("2013/10/12")
instrument1.putCallType = callOption

Dim spot As Double: spot = 110
MsgBox instrument1.calculatePayoff(spot)
End Sub
