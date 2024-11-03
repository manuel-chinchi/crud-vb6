Attribute VB_Name = "modCategoryHelper"
'TODO implementar algo mas flexible para parsear los objetos. Revisar el obj
' Dictionary si puede servir

Option Explicit

Dim i As Integer

Public Function NewCategory(ParamArray varParam() As Variant) As Object
    Dim Category As clsCategory
    Set Category = New clsCategory
    
    With Category
        .mId = varParam(0)
        .mName = varParam(1)
        .mStatus = CBool(varParam(2))
        Set .mArticles = varParam(3)
        .mCreateAt = CDate(varParam(4))
        .mUpateAt = CDate(varParam(5))
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

Public Function ConvertToRecordset(arr As Collection) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim obj As clsCategory
    
    rs.Fields.Append "mId", adInteger
    rs.Fields.Append "mName", adVarChar, 255
    rs.Fields.Append "mArticlesCount", adInteger
    
    rs.Open
    Dim iArticlesCount As Integer
    For Each obj In arr
        rs.AddNew
        rs.Fields("mId").Value = obj.mId
        rs.Fields("mName").Value = obj.mName
        If obj.mArticles Is Nothing Then
            iArticlesCount = 0
        Else
            iArticlesCount = obj.mArticles.Count
        End If
        'If obj.mArticles Is Nothing Then
        '    Set rs.Fields("mArticlesCount").Value = obj.mArticles.Count
        'Else
        '    rs.Fields("mArticlesCount").Value = 0
        'End If
        rs.Fields("mArticlesCount").Value = iArticlesCount
        rs.Update
    Next obj
    
    Set ConvertToRecordset = rs
End Function
