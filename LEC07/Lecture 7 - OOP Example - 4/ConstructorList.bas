Attribute VB_Name = "ConstructorList"
Option Explicit
Option Base 1

Public Function newPlainVanilla(strike As Double, _
                                                         expiryDate As Date, _
                                                         putCallType As OptionType) As PlainVanilla
                                                        
Dim newOption As New PlainVanilla
newOption.strike = strike
newOption.expiryDate = expiryDate
newOption.putCallType = putCallType
Set newPlainVanilla = newOption
End Function

