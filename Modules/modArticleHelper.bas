Attribute VB_Name = "modArticleHelper"
'TODO implementar algo mas flexible para parsear los objetos. Revisar el obj
' Dictionary

Public Function NewArticle(ParamArray varParams() As Variant) As Object
    Dim Article As clsArticle
    Set Article = New clsArticle
    
    With Article
        .mId = varParams(0)
        .mName = varParams(1)
        .mDetails = varParams(2)
        .mCreateAt = CDate(varParams(3))
        .mUpdateAt = CDate(varParams(4))
        '.mCategoryName = varParams(5)
        If .mCategory Is Nothing Then
            Set .mCategory = New clsCategory
        End If
        Set .mCategory = varParams(6)
        '.mCategoryId = varParams(7)
    End With
    
    Set NewArticle = Article
End Function

' needs "Microsoft ActiveX Data Objects 2.6 Library" reference
Public Function ConvertToRecordset(arr As Collection) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim obj As clsArticle
    
    rs.Fields.Append "mId", adInteger
    rs.Fields.Append "mName", adVarChar, 255
    rs.Fields.Append "mDetails", adVarChar, 255
    rs.Fields.Append "mCategoryName", adVarChar, 255
    
    rs.Open
    
    For Each obj In arr
        rs.AddNew
        rs.Fields("mId").Value = obj.mId
        rs.Fields("mName").Value = obj.mName
        rs.Fields("mDetails").Value = obj.mDetails
        'rs.Fields("mCategoryName").Value = obj.mCategoryName
        rs.Fields("mCategoryName").Value = obj.mCategory.mName
        rs.Update
    Next obj
    
    Set ConvertToRecordset = rs
End Function
