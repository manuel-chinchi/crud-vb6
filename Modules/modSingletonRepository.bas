Attribute VB_Name = "modSingletonRepository"
Option Explicit

Private mInstanceAR As clsArticleRepository
Private mInstanceCR As clsCategoryRepository

Public Function GetArticleRepository() As clsArticleRepository
    If mInstanceAR Is Nothing Then
        Set mInstanceAR = New clsArticleRepository
    End If
    
    Set GetArticleRepository = mInstanceAR
End Function

Public Function GetCategoryRepository() As clsCategoryRepository
    If mInstanceCR Is Nothing Then
        Set mInstanceCR = New clsCategoryRepository
    End If
    
    Set GetCategoryRepository = mInstanceCR
End Function
