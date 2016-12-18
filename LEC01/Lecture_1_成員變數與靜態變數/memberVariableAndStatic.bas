Attribute VB_Name = "memberVariableAndStatic"
Option Explicit
'Version 1
Function emaCalculator(priceArray As Range, alpha As Variant) As Variant

Dim dominator As Variant                                         '¥[Åv¥­§¡¥[Á`©Ò°£¤Wªº¤À¥À
Dim onePrice As Range                                            '¬Y¤Ñ»ù®æªºÀx¦s®æ
Dim weight As Variant                                                                       '¸Ó»ù®æ¤§Åv­«
Dim counts As Integer: counts = priceArray.Count                 '²{¦bÅv­«ªº alpha ¤§¦¸¤è

For Each onePrice In priceArray
    counts = counts - 1
    weight = alpha ^ counts
    emaCalculator = emaCalculator + onePrice.Value * weight
    dominator = dominator + weight
Next onePrice

emaCalculator = emaCalculator / dominator
End Function
