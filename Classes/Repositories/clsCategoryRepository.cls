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
    Dim cCategories As New Collection
    
    mConnection.OpenDB
    mConnection.ExecuteQuery modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_get_categories.sql")
    
    Dim dCategory As Dictionary
    Dim cItemsC As New Collection
    
    Set cItemsC = mConnection.Items
    
    For Each dCategory In cItemsC
        Dim oCategory As New clsCategory
        
        mConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_get_articles_of_category.sql")
        mConnection.AddWithValue "@CategoryId", dCategory("c.Id")
        mConnection.ExecuteNonQuery
        
        Dim dArticle As Dictionary
        Dim cItemsA As New Collection
        
        Set cItemsA = mConnection.Items
        
        Dim cArticles As New Collection
        
        For Each dArticle In cItemsA
            Dim oArticle As clsArticle
            
            Set oArticle = modArticleHelper.Create( _
                dArticle("Id"), _
                dArticle("Name"), _
                dArticle("Details"), _
                Nothing, _
                dArticle("CreateAt"), _
                dArticle("UpdateAt") _
            )
            
            cArticles.Add oArticle
        Next dArticle
        
        Set oCategory = modCategoryHelper.Create( _
            dCategory("c.Id"), _
            dCategory("c.Name"), _
            dCategory("c.Status"), _
            dCategory("c.CreateAt"), _
            dCategory("c.UpdateAt"), _
            cArticles _
        )
        
        cCategories.Add oCategory
        Set oCategory = Nothing
        Set cArticles = Nothing
    Next dCategory
    
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
