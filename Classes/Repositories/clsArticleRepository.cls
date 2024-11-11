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
Dim mConnection As clsSQLiteConnection
Dim i As Integer

Private Sub Class_Initialize()
    Set mArticles = New Collection
    Set mConnection = New clsSQLiteConnection
    mConnection.filePath = App.Path & "\Data\database.db"
End Sub

Private Sub Class_Terminate()
    Set mArticles = Nothing
    Set mConnection = Nothing
End Sub


Public Function GetArticles() As Collection
On Error GoTo Catch
    Dim cArticles As New Collection
    
    mConnection.OpenDB
    mConnection.ExecuteQuery modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_get_articles.sql")
    
    Dim cItems As Collection
    Set cItems = mConnection.Items
            
    Dim oDict As Dictionary
    If Not cItems Is Nothing Then
        For Each oDict In cItems
            mConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_get_category_of_article.sql")
            mConnection.AddWithValue "@CategoryId", oDict("CategoryId")
            mConnection.ExecuteNonQuery
            
            Dim cItems2 As Collection
            Set cItems2 = mConnection.Items
            
            Dim oDict2 As Dictionary
            For Each oDict2 In cItems2
                Dim oCategory As clsCategory
                Set oCategory = modCategoryHelper.Create( _
                    oDict2("Id"), _
                    oDict2("Name"), _
                    oDict2("State"), _
                    oDict2("CreateAt"), _
                    oDict2("UpdateAt"), _
                    Nothing _
                )
            Next oDict2
            
            Dim oArticle As New clsArticle
            Set oArticle = modArticleHelper.Create( _
                oDict("Id"), _
                oDict("Name"), _
                oDict("Details"), _
                oCategory, _
                oDict("CreateAt"), _
                oDict("UpdateAt") _
            )
            cArticles.Add oArticle
            Set oArticle = Nothing
            Set oCategory = Nothing
        Next oDict
    End If
    
    mConnection.CloseDB
    Set GetArticles = cArticles
    Exit Function
    
Catch:
    Err.Raise vbObjectError + 513, , Err.Number & " " & Err.Description & " on GetCategories"
End Function

Public Sub CreateArticle(obj As clsArticle)
    mConnection.OpenDB
    mConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_create_article.sql")
    mConnection.AddWithValue "@Name", obj.mName
    mConnection.AddWithValue "@Details", obj.mDetails
    mConnection.AddWithValue "@CreateAt", GetDateTimeAsString
    mConnection.AddWithValue "@CategoryId", obj.mCategory.mId
    mConnection.ExecuteNonQuery
    mConnection.CloseDB
End Sub

Public Sub UpdateArticle(obj As clsArticle)
    mConnection.OpenDB
    mConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_update_article.sql")
    mConnection.AddWithValue "@Id", obj.mId
    mConnection.AddWithValue "@Name", obj.mName
    mConnection.AddWithValue "@Details", obj.mDetails
    mConnection.AddWithValue "@CategoryId", obj.mCategory.mId
    mConnection.AddWithValue "@UpdateAt", GetDateTimeAsString
    mConnection.ExecuteNonQuery
    mConnection.CloseDB
End Sub

Public Sub DeleteArticle(Id As Long)
    mConnection.OpenDB
    mConnection.Query = modIOHelper.GetStringFromFile(App.Path + "\Scripts\sp_delete_article.sql")
    mConnection.AddWithValue "@Id", Id
    mConnection.ExecuteNonQuery
    mConnection.CloseDB
End Sub

Public Function SearchArticle(sSearch As String) As Collection
On Error GoTo Catch
    Dim cArticles As New Collection
    
    mConnection.OpenDB
    mConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_search_article.sql")
    mConnection.AddWithValue "@Search", sSearch
    mConnection.ExecuteNonQuery
    
    Dim oDict As Dictionary
    Dim cItems As Collection
    Set cItems = mConnection.Items
    
    If Not cItems Is Nothing Then
        For Each oDict In cItems
            mConnection.Query = modIOHelper.GetStringFromFile(App.Path & "\Scripts\sp_get_category_of_article.sql")
            mConnection.AddWithValue "@CategoryId", oDict("CategoryId")
            mConnection.ExecuteNonQuery
            
            Dim oDict2 As Dictionary
            Dim cItems2 As Collection
            Set cItems2 = mConnection.Items
            
            For Each oDict2 In cItems2
                Dim oCategory As clsCategory
                Set oCategory = modCategoryHelper.Create( _
                    oDict2("Id"), _
                    oDict2("Name"), _
                    oDict2("Status"), _
                    oDict2("CreateAt"), _
                    oDict2("UpdateAt"), _
                    Nothing _
                )
            Next oDict2
            
            Dim oArticle As New clsArticle
            Set oArticle = modArticleHelper.Create( _
                oDict("Id"), _
                oDict("Name"), _
                oDict("Details"), _
                oCategory, _
                oDict("CreateAt"), _
                oDict("UpdateAt") _
            )
            
            cArticles.Add oArticle
            Set oArticle = Nothing
            Set oCategory = Nothing
        Next oDict
    End If
    
    mConnection.CloseDB
    Set SearchArticle = cArticles
    Exit Function
    
Catch:
    Err.Raise vbObjectError + 513, , Err.Number & " " & Err.Description & " on SearchArticle2"
End Function

Private Function GetDateTimeAsString() As String
    GetDateTimeAsString = Format(Now, "dd-mm-yyyy  hh:mm:ss")
End Function
