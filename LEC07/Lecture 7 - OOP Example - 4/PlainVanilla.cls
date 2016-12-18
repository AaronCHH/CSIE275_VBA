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

' putCallType_ = 1 (call)
' putCallType_ = -1 (put)
Private putCallType_ As Integer
' ========================================================================= '
Public expiryDate As Date

' ========================================================================= '
Public Property Let putCallType(ByVal value As OptionType)

Select Case value
    Case callOption
        putCallType_ = 1
    Case putOption
        putCallType_ = -1
End Select
    
End Property
' ========================================================================= '
Public Property Get putCallType() As OptionType
putCallType = getOptionType()
End Property

Private Function getOptionType() As OptionType
Select Case putCallType_
    Case 1
        getOptionType = callOption
    Case -1
        getOptionType = putOption
End Select
End Function

' ========================================================================= '
Public Property Let strike(ByVal value As Double)

If value > 0 Then
    strike_ = value
Else
    Err.Raise Number:=10001, Source:="strike", Description:="履約價須大於零"
End If
    
End Property
' ========================================================================= '
Public Property Get strike() As Double
strike = strike_
End Property

' ========================================================================= '
Public Function calculatePayoff(spot As TimeSeries) As TimeSeries
Dim nDays As Integer: nDays = UBound(spot.dates)
ReDim payoff(nDays) As Double
Dim exercisePayoff As Double
Dim i As Integer

For i = 1 To nDays
    If spot.dates(i) = expiryDate Then
        exercisePayoff = putCallType_ * (spot.values(i) - strike_)
        payoff(i) = IIf(exercisePayoff > 0, exercisePayoff, 0)
        Exit For
    End If
Next i

calculatePayoff = fromArray(spot.dates, payoff, "payoff")
End Function
' ========================================================================= '
Public Function clone() As PlainVanilla

Dim cloneOption As New PlainVanilla
cloneOption.expiryDate = expiryDate
cloneOption.strike = strike_
cloneOption.putCallType = getOptionType()
Set clone = cloneOption
End Function



















