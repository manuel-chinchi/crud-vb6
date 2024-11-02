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
Dim DBConnection As clsSQLiteConnection
Dim PathDB As String
Dim i As Integer

Private Sub Class_Initialize() 'ctor
    Set Articles = New Collection
    Set DBConnection = New clsSQLiteConnection
    DBConnection.FilePath = App.Path & "\Data\database.db"
End Sub

Private Sub Class_Terminate()
    Set Articles = Nothing
    Set DBConnection = Nothing
End Sub

Public Function GetArticles() As Collection
    Dim vArticles As Variant
    
    DBConnection.OpenDB
    Dim sQuery As String
    sQuery = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_get_articles.sql")
    DBConnection.ExecuteQuery sQuery
    vArticles = DBConnection.Cursor
    DBConnection.CloseDB
    
    If IsEmpty(vArticles) Then
        Set GetArticles = Nothing
    Else
        Dim oArticle As New clsArticle
        Dim cArticles As New Collection
        
        For i = 0 To UBound(vArticles)
            Set oArticle = modArticleHelper.NewArticle( _
                    vArticles(i, 0), _
                    vArticles(i, 1), _
                    vArticles(i, 2), _
                    vArticles(i, 5))
            cArticles.Add oArticle
        Next
    End If

    Set GetArticles = cArticles
End Function

Public Sub CreateArticle(obj As clsArticle)
    
    DBConnection.OpenDB
    DBConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_create_article.sql")
    DBConnection.AddWithValue "@Name", obj.mName
    DBConnection.AddWithValue "@Details", obj.mDetails
    DBConnection.AddWithValue "@CreateAt", GetDateTimeAsString
    'DBConnection.AddWithValue "@UpdateAt", Null
    DBConnection.AddWithValue "@CategoryId", obj.mCategoryId
    DBConnection.ExecuteNonQuery
    DBConnection.CloseDB
End Sub

Public Sub UpdateArticle(obj As clsArticle)
    DBConnection.OpenDB
    DBConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_update_article.sql")
    DBConnection.AddWithValue "@Id", obj.mId
    DBConnection.AddWithValue "@Name", obj.mName
    DBConnection.AddWithValue "@Details", obj.mDetails
    DBConnection.AddWithValue "@CategoryId", obj.mCategoryId
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
    Dim vArticles As Variant
    
    DBConnection.OpenDB
    DBConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_search_article.sql")
    DBConnection.AddWithValue "@Search", sSearch
    DBConnection.ExecuteNonQuery
    vArticles = DBConnection.Cursor
    DBConnection.CloseDB
    
    If IsEmpty(vArticles) Then
        Set SearchArticle = Nothing
    Else
        Dim oArticle As New clsArticle
        Dim cArticles As New Collection
        
        For i = 0 To UBound(vArticles)
            Set oArticle = modArticleHelper.NewArticle( _
                    vArticles(i, 0), _
                    vArticles(i, 1), _
                    vArticles(i, 2), _
                    vArticles(i, 5))
            cArticles.Add oArticle
        Next
    End If
    
    Set SearchArticle = cArticles
End Function

Function GetDateTimeAsString() As String
    GetDateTimeAsString = Format(Now, "dd-mm-yyyy  hh:mm:ss")
End Function
