VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticleRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Articles As Collection

Private Sub Class_Initialize() 'ctor
    Set Articles = New Collection
    
    Articles.Add modArticleHelper.NewArticle(1, "buzo", "5xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(2, "remera", "20xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(300, "jean", "40xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(4, "medias", "400xU", "calzado")
    
End Sub

Private Sub Class_Terminate()
    Set Articles = Nothing
End Sub

Public Function GetArticles() As Collection
    Set GetArticles = Articles
End Function

Public Sub CreateArticle(obj As clsArticle)
    Articles.Add obj
End Sub

Public Sub UpdateArticle(obj As clsArticle)
    Dim oArticle As clsArticle
    
    For Each oArticle In Articles
        If oArticle.mId = obj.mId Then
            With oArticle
                .mName = obj.mName
                .mDetails = obj.mDetails
                .mCategoryName = obj.mCategoryName
            End With
        End If
    Next
End Sub

Public Sub DeleteArticle(Id As Long)
    Dim oArticle As clsArticle
    Set oArticle = New clsArticle
    
    Dim iIndex As Integer
    For Each oArticle In Articles
        If oArticle.mId = Id Then
            Exit For
        End If
        iIndex = iIndex + 1
    Next
    
    Articles.Remove (iIndex + 1)
End Sub

Public Function SearchArticle(sSearch As String) As Collection
    Dim oArticle As clsArticle
    Dim arrArticlesFilter As Collection
    Set arrArticlesFilter = New Collection
    
    For Each oArticle In Articles
        If InStr(1, LCase(oArticle.ToString()), LCase(sSearch)) <> 0 Then
            arrArticlesFilter.Add oArticle
        End If
    Next
    Set SearchArticle = arrArticlesFilter
End Function
