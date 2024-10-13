Attribute VB_Name = "modCategoryHelper"
Public Function NewCategory(ParamArray varParam() As Variant) As Object
    Dim Category As clsCategory
    Set Category = New clsCategory
    
    With Category
        .mId = varParam(0)
        .mName = varParam(1)
        Set .mArticlesRelated = varParam(2)
    End With
    
    Set NewCategory = Category
End Function
