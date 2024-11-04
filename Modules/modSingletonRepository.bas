Attribute VB_Name = "modSingletonRepository"
Option Explicit

Private mInstanceArticleRepository As clsArticleRepository
Private mInstanceCategoryRepository As clsCategoryRepository

Public Function GetArticleRepository() As clsArticleRepository
    If mInstanceArticleRepository Is Nothing Then
        Set mInstanceArticleRepository = New clsArticleRepository
    End If
    
    Set GetArticleRepository = mInstanceArticleRepository
End Function

Public Function GetCategoryRepository() As clsCategoryRepository
    If mInstanceCategoryRepository Is Nothing Then
        Set mInstanceCategoryRepository = New clsCategoryRepository
    End If
    
    Set GetCategoryRepository = mInstanceCategoryRepository
End Function
