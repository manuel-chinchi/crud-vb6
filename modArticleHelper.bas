Attribute VB_Name = "modArticleHelper"
Public Function NewArticle(ParamArray varParams() As Variant) As Object
    Dim Article As clsArticle
    Set Article = New clsArticle
    
    With Article
        .mId = varParams(0)
        .mName = varParams(1)
        .mDetails = varParams(2)
        .mCategoryName = varParams(3)
    End With
    
    Set NewArticle = Article
End Function

