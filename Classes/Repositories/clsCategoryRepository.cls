VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCategoryRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Categories As Collection

Private Sub Class_Initialize()
    Set Categories = New Collection
    
    Dim Articles As Collection
    Set Articles = New Collection
    
    Articles.Add modArticleHelper.NewArticle(1, "buzo", "5xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(2, "remera", "20xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(300, "jean", "40xU", "indumentaria")
    
    Categories.Add modCategoryHelper.NewCategory(1, "indumentaria", Articles)
    Categories.Add modCategoryHelper.NewCategory(2, "remeras", Nothing)
    Categories.Add modCategoryHelper.NewCategory(3, "pantalones", Nothing)
    
    Set Articles = New Collection
    Articles.Add modArticleHelper.NewArticle(4, "medias", "400xU", "calzado")
    
    Categories.Add modCategoryHelper.NewCategory(4, "calzado", Articles)
End Sub

Private Sub Class_Terminate()
    Set Categories = Nothing
End Sub

Public Function GetCategories() As Collection
    Set GetCategories = Categories
End Function

Public Sub CreateCategory(obj As clsCategory)
    Categories.Add obj
End Sub

'Public Sub UpdateCategory(obj As clsCategory)
'    Dim oCategory As clsCategory
'
'    For Each oCategory In Categories
'        If oCategory.mId = obj.mId Then
'            With oCategory
'                .mName = obj.mName
'                Set .mArticlesRelated = obj.mArticlesRelated
'            End With
'        End If
'    Next
'End Sub

Public Sub DeleteCategory(Id As Integer)
    Dim oCategory As clsCategory
    Set oCategory = New clsCategory
    
    Dim iIndex As Integer
    For Each oCategory In Categories
        If oCategory.mId = Id Then
            Exit For
        End If
        iIndex = iIndex + 1
    Next
    
    Categories.Remove (iIndex + 1)
End Sub