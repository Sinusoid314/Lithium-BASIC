VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public itemCount As Long
Private items() As Variant


Public Sub Add(ByVal newItem As Variant, Optional ByVal before As Integer)

Dim tmp() As Variant

ReDim tmp(itemCount)

If before = Empty Or itemCount = 0 Then
    For I = 1 To itemCount
        If IsObject(items(I)) Then Set tmp(I) = items(I) Else tmp(I) = items(I)
    Next I
    ReDim items(itemCount + 1)
    For I = 1 To itemCount
        If IsObject(tmp(I)) Then Set items(I) = tmp(I) Else items(I) = tmp(I)
    Next I
    If IsObject(newItem) Then Set items(itemCount + 1) = newItem Else items(itemCount + 1) = newItem
ElseIf before > 0 And before <= itemCount Then
    For I = 1 To itemCount
        If IsObject(items(I)) Then Set tmp(I) = items(I) Else tmp(I) = items(I)
    Next I
    ReDim items(itemCount + 1)
    For I = 1 To before - 1
        If IsObject(tmp(I)) Then Set items(I) = tmp(I) Else items(I) = tmp(I)
    Next I
    If IsObject(newItem) Then Set items(before) = newItem Else items(before) = newItem
    For I = before + 1 To itemCount + 1
        If IsObject(tmp(I - 1)) Then Set items(I) = tmp(I - 1) Else items(I) = tmp(I - 1)
    Next I
End If

itemCount = itemCount + 1

End Sub

Public Sub Clear()


ReDim items(0)

itemCount = 0


End Sub

Public Property Get Item(ByVal itemIdx As Integer) As Variant

If itemIdx > 0 And itemIdx <= itemCount Then
    If IsObject(items(itemIdx)) Then Set Item = items(itemIdx) Else Item = items(itemIdx)
Else
    MsgBox "ArrayClass error: Index '" & itemIdx & "' is out of bounds"
    a = 1
End If

End Property
Public Property Let Item(ByVal itemIdx As Integer, newItem As Variant)

If itemIdx > 0 And itemIdx <= itemCount Then
    If IsObject(newItem) Then Set items(itemIdx) = newItem Else items(itemIdx) = newItem
Else
    MsgBox "ArrayClass error: Index '" & itemIdx & "' is out of bounds"
    a = 1
End If

End Property
Public Sub Remove(ByVal itemIdx As Integer)

If itemCount = 0 Or itemIdx <= 0 Or itemIdx > itemCount Then Exit Sub

Dim tmp() As Variant

ReDim tmp(itemCount)

For I = 1 To itemCount
    If IsObject(items(I)) Then Set tmp(I) = items(I) Else tmp(I) = items(I)
Next I
ReDim items(itemCount - 1)
For I = 1 To itemIdx - 1
    If IsObject(tmp(I)) Then Set items(I) = tmp(I) Else items(I) = tmp(I)
Next I
For I = itemCount To itemIdx + 1 Step -1
    If IsObject(tmp(I)) Then Set items(I - 1) = tmp(I) Else items(I - 1) = tmp(I)
Next I

itemCount = itemCount - 1

End Sub

Private Sub Class_Initialize()

itemCount = 0

End Sub


