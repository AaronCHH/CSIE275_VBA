VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlainVanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private strike_ As Double
Public expiryDate As Date
Public putCallType As OptionType

Public Property Let strike(ByVal value As Double)

If value > 0 Then
    strike_ = value
Else
    Err.Raise Number:=10001, Source:="strike", Description:="履約價須大於零"
End If
    
End Property

Public Property Get strike() As Double
strike = strike_
End Property

Public Function calculatePayoff(spot As Double) As Double
Select Case putCallType
    Case callOption
        calculatePayoff = IIf(spot > strike, spot - strike, 0)
    Case putOption
        calculatePayoff = IIf(spot < strike, strike - spot, 0)
End Select
End Function



















