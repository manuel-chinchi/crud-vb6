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
        rs.Fields("mCategoryName").Value = obj.mCategoryName
        rs.Update
    Next obj
    
    Set ConvertToRecordset = rs
End Function
