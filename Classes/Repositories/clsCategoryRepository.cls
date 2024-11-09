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

Dim mCategories As Collection
Dim mConnection As clsSQLiteConnection
Dim mArticleRepository As clsArticleRepository
Dim i, j As Integer

Private Sub Class_Initialize()
    Set mCategories = New Collection
    Set mConnection = New clsSQLiteConnection
    Set mArticleRepository = New clsArticleRepository
    mConnection.filePath = App.Path + "\Data\database.db"
End Sub

Private Sub Class_Terminate()
    Set mCategories = Nothing
    Set mConnection = Nothing
End Sub

Public Function GetCategories() As Collection
    'NOTA: lanzar error en mConnection si no se cargo el parametro .FilePath
    'Set GetCategories = mCategories
    
    Dim vCategories As Variant
    Dim cCategories As New Collection
    
    mConnection.OpenDB
    mConnection.ExecuteQuery modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_get_categories.sql")
    vCategories = mConnection.Cursor

    Dim oCategory As New clsCategory
    For i = 0 To UBound(vCategories)

        Dim oArticle As clsArticle
        Dim vArticles As Variant
        Dim cArticles As New Collection
        
        mConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_get_articles_of_category.sql")
        mConnection.AddWithValue "@CategoryId", vCategories(i, 0)
        mConnection.ExecuteNonQuery
        vArticles = mConnection.Cursor
        
        If Not IsEmpty(vArticles) Then
            For j = 0 To UBound(vArticles)
                Set oArticle = modArticleHelper.NewArticle( _
                    vArticles(j, 0), _
                    vArticles(j, 1), _
                    vArticles(j, 2), _
                    vArticles(j, 3), _
                    vArticles(j, 4), _
                    vArticles(j, 5), _
                    Nothing _
                    )
                cArticles.Add oArticle
            Next
        End If

        Set oCategory = modCategoryHelper.NewCategory( _
            vCategories(i, 0), _
            vCategories(i, 1), _
            vCategories(i, 2), _
            cArticles, _
            vCategories(i, 3), _
            vCategories(i, 4) _
            )
        
        cCategories.Add oCategory
        Set cArticles = New Collection
    Next
    
    mConnection.CloseDB
    Set GetCategories = cCategories
End Function

Public Sub CreateCategory(obj As clsCategory)
    mConnection.OpenDB
    mConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_create_category.sql")
    mConnection.AddWithValue "@Name", obj.mName
    mConnection.AddWithValue "@CreateAt", Format(Now, "dd-mm-yyyy hh:mm:ss")
    mConnection.ExecuteNonQuery
    mConnection.CloseDB
End Sub

Public Sub DeleteCategory(Id As Integer)
    mConnection.OpenDB
    mConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_delete_category.sql")
    mConnection.AddWithValue "@Id", Id
    mConnection.ExecuteNonQuery
    mConnection.CloseDB
End Sub
