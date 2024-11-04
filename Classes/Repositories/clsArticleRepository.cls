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

Dim mArticles As Collection
Dim DBConnection As clsSQLiteConnection
Dim PathDB As String
Dim i As Integer

Private Sub Class_Initialize() 'ctor
    Set mArticles = New Collection
    Set DBConnection = New clsSQLiteConnection
    DBConnection.filePath = App.Path & "\Data\database.db"
End Sub

Private Sub Class_Terminate()
    Set mArticles = Nothing
    Set DBConnection = Nothing
End Sub

Public Function GetArticles() As Collection
    Dim vmArticles As Variant
    
    DBConnection.OpenDB
    Dim sQuery As String
    sQuery = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_get_articles.sql")
    DBConnection.ExecuteQuery sQuery
    vmArticles = DBConnection.Cursor
    
    If IsEmpty(vmArticles) Then
        Set GetArticles = Nothing
    Else
        Dim oArticle As New clsArticle
        Dim cmArticles As New Collection
        
        For i = 0 To UBound(vmArticles)
            
            Dim vCategory As Variant
            Dim oCategory As clsCategory
            
            DBConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_get_category_of_article.sql")
            DBConnection.AddWithValue "@CategoryId", vmArticles(i, 6)
            DBConnection.ExecuteNonQuery
            vCategory = DBConnection.Cursor
            
            Set oCategory = modCategoryHelper.NewCategory( _
                vCategory(0, 0), _
                vCategory(0, 1), _
                vCategory(0, 2), _
                Nothing, _
                vCategory(0, 3), _
                vCategory(0, 4) _
            )
        
            Set oArticle = modArticleHelper.NewArticle( _
                    vmArticles(i, 0), _
                    vmArticles(i, 1), _
                    vmArticles(i, 2), _
                    vmArticles(i, 3), _
                    vmArticles(i, 4), _
                    vmArticles(i, 5), _
                    oCategory, _
                    vmArticles(i, 6) _
                    )
            cmArticles.Add oArticle
        Next
    End If
    
    DBConnection.CloseDB
    Set GetArticles = cmArticles
End Function

Public Sub CreateArticle(obj As clsArticle)
    
    DBConnection.OpenDB
    DBConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_create_article.sql")
    DBConnection.AddWithValue "@Name", obj.mName
    DBConnection.AddWithValue "@Details", obj.mDetails
    DBConnection.AddWithValue "@CreateAt", GetDateTimeAsString
    'DBConnection.AddWithValue "@UpdateAt", Null
    'DBConnection.AddWithValue "@CategoryId", obj.mCategoryId
    DBConnection.AddWithValue "@CategoryId", obj.mCategory.mId
    DBConnection.ExecuteNonQuery
    DBConnection.CloseDB
End Sub

Public Sub UpdateArticle(obj As clsArticle)
    DBConnection.OpenDB
    DBConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_update_article.sql")
    DBConnection.AddWithValue "@Id", obj.mId
    DBConnection.AddWithValue "@Name", obj.mName
    DBConnection.AddWithValue "@Details", obj.mDetails
    'DBConnection.AddWithValue "@CategoryId", obj.mCategoryId
    DBConnection.AddWithValue "@CategoryId", obj.mCategory.mId
    DBConnection.AddWithValue "@UpdateAt", GetDateTimeAsString
    DBConnection.ExecuteNonQuery
    DBConnection.CloseDB
End Sub

Public Sub DeleteArticle(Id As Long)
    DBConnection.OpenDB
    DBConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_delete_article.sql")
    DBConnection.AddWithValue "@Id", Id
    DBConnection.ExecuteNonQuery
    DBConnection.CloseDB
End Sub

Public Function SearchArticle(sSearch As String) As Collection
    Dim vmArticles As Variant
    
    DBConnection.OpenDB
    DBConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_search_article.sql")
    DBConnection.AddWithValue "@Search", sSearch
    DBConnection.ExecuteNonQuery
    vmArticles = DBConnection.Cursor

    If IsEmpty(vmArticles) Then
        Set SearchArticle = Nothing
    Else
        Dim oArticle As New clsArticle
        Dim cmArticles As New Collection
        
        For i = 0 To UBound(vmArticles)
        
            Dim vCategory As Variant
            Dim oCategory As clsCategory
            
            DBConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_get_category_of_article.sql")
            DBConnection.AddWithValue "@CategoryId", vmArticles(i, 6)
            DBConnection.ExecuteNonQuery
            vCategory = DBConnection.Cursor
            
            Set oCategory = modCategoryHelper.NewCategory( _
                vCategory(0, 0), _
                vCategory(0, 1), _
                vCategory(0, 2), _
                Nothing, _
                vCategory(0, 3), _
                vCategory(0, 4) _
            )
        
            Set oArticle = modArticleHelper.NewArticle( _
                    vmArticles(i, 0), _
                    vmArticles(i, 1), _
                    vmArticles(i, 2), _
                    vmArticles(i, 3), _
                    vmArticles(i, 4), _
                    vmArticles(i, 5), _
                    oCategory, _
                    vmArticles(i, 6) _
                )
                    
            cmArticles.Add oArticle
        Next
    End If
    
    DBConnection.CloseDB
    
    Set SearchArticle = cmArticles
End Function

Function GetDateTimeAsString() As String
    GetDateTimeAsString = Format(Now, "dd-mm-yyyy  hh:mm:ss")
End Function
