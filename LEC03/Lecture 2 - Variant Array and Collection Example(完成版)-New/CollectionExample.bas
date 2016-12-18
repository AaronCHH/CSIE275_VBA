Attribute VB_Name = "CollectionExample"
Option Explicit
Option Base 1

Sub dataToCollection()

Dim tw2311 As Collection: Set tw2311 = New Collection
'�άO Dim tw2311 As New Collection

'���i�J�{�Ȥu�@��A��J���
Worksheets("���զX�{��").Activate

'Item : ���e��
'Key : ���ҦW��
tw2311.Add Item:=Range("B1").Value, key:="presentValueDate"
'�{��
tw2311.Add Item:=Range("B5").Value, key:="presentValue", _
                       before:="presentValueDate"
MsgBox hasKey(tw2311, "ABC")
Set tw2311 = setItem(tw2311, "ABC", tw2311("presentValueDate") + 1)
MsgBox hasKey(tw2311, "ABC")
End Sub

Function hasKey(col As Collection, key As String) As Boolean
Dim temp As Variant
On Error GoTo doNotHaveKey
    temp = col(key)
    hasKey = True
    Exit Function
doNotHaveKey:
    hasKey = False
End Function

Function setItem(col As Collection, key As String, newItem As Variant) As Collection

If hasKey(col, key) Then col.Remove (key)
    
col.Add Item:=newItem, key:=key
Set setItem = col
    
End Function

Sub collectionExample()

Dim personalInfo As New Collection

personalInfo.Add Item:="�B�۱j", key:="name"
personalInfo.Add Item:=50, key:="age"

End Sub

Sub forEachExample()

Dim x As New Collection
  x.Add Item:=100, key:="value"
  x.Add Item:="AAA", key:="string"

Dim elem As Variant
  For Each elem In x
      Debug.Print elem
  Next elem
End Sub












