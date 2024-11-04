VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public mId As Integer
Public mName As String
Public mDetails As String
Public mCategory As clsCategory
Public mCreateAt As Date
Public mUpdateAt As Date

Public Function ToString() As String
    ToString = _
        "<Id=" & Me.mId & _
        ", Name=" & Me.mName & _
        ", Details=" & Me.mDetails & _
        ", CreateAt=" & Me.mCreateAt & _
        ", UpdateAt=" & Me.mUpdateAt & _
        ", Category= <mId=" & Me.mCategory.mId & ">>"
End Function
