Attribute VB_Name = "example_change"
Option Explicit
'�����ܼ�
Dim memberVar As Integer

Sub counts()
'�ϰ��ܼ� : �s�򪺮ɶ��P�l�{�ǰ���ɶ��@�ˤ[
Dim localVar As Integer

localVar = localVar + 1
memberVar = memberVar + 1

MsgBox "Member:" & memberVar & Chr(10) & _
                 "Local:" & localVar
End Sub
