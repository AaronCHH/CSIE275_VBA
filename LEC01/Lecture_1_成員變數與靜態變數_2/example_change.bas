Attribute VB_Name = "example_change"
Option Explicit
'成員變數
Dim memberVar As Integer

Sub counts()
'區域變數 : 存續的時間與子程序執行時間一樣久
Dim localVar As Integer

localVar = localVar + 1
memberVar = memberVar + 1

MsgBox "Member:" & memberVar & Chr(10) & _
                 "Local:" & localVar
End Sub
