Attribute VB_Name = "modCategoryHelper"
Option Explicit

Dim i As Integer

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

Public Function ConvertToVariant(arr As Collection) As Variant
    Dim oCategory As clsCategory
    Dim vArray As Variant
    
    i = 1
    ReDim vArray(i To arr.Count)
    
    For Each oCategory In arr
        vArray(i) = oCategory.mName
        i = i + 1
    Next
    ConvertToVariant = vArray
End Function
