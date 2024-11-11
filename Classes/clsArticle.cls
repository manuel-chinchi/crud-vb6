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

Public Static Function Create(Id As Integer, _
                        Name As String, _
                        Details As String, _
                        Category As clsCategory, _
                        CreateAt As Date, _
                        UpdateAt As Date) As clsArticle
    Dim obj As New clsArticle
        
    With obj
        .mId = Id
        .mName = Name
        .mDetails = Details
        Set mCategory = Category
        .mCreateAt = CreateAt
        .mUpdateAt = UpdateAt
    End With
    
    Set CreateArticle = obj
End Function

Public Function ToString() As String
    Dim iCategoryId As Integer
    If Me.mCategory Is Nothing Then
        iCategoryId = -1
    Else
        iCategoryId = Me.mCategory.mId
    End If
    ToString = _
        "<Id=" & Me.mId & _
        ", Name=" & Me.mName & _
        ", Details=" & Me.mDetails & _
        ", CreateAt=" & Me.mCreateAt & _
        ", UpdateAt=" & Me.mUpdateAt & _
        ", Category= <mId=" & iCategoryId & ">>"
End Function
