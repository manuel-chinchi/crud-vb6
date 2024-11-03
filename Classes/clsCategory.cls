VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public mId As Integer
Public mName As String
Public mStatus As Boolean
Public mCreateAt As Date
Public mUpateAt As Date
Public mArticles As Collection

Public Property Get ToString() As String
    If Not mArticles Is Nothing Then
        ToString = mId & "," & mName & "," & Str(mArticles.Count)
    Else
        ToString = mId & "," & mName & "," & 0
    End If
End Property
