VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLiteConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
' File: cSQLiteConnection.cls
' Author: Manuel Chinchi (Argentina)
' Source:
' Licence: GPL3
'
'
' Members:
' - FilePath:       Path to database file.
' - Query:          Sentence to executed.
' - Handle:         Database instance handler.
' - Headers:        Headers retrieved from the query.
' - Cursor:         Cursor containts the result of the query.
' - CursorString:   Cursor containts the query result in string format.
' - Error:          Errors when executing the query.
' - Rows:           Numbers of rows.
' - Columns:        Numbers of columns.
' - RowLimit:       Limit row index.
' - ColumnLimit:    Limit column index.
'
' Methods:
' - OpenDB([sFilePath]):             Open the database.
' - ExecuteQuery([sQuery],[fIncludeHeaders]):       Executes specified query in 'Query' member.
' - ExecuteNonQuery([fIncludeHeaders]):    Executes at especific query type INSERT, UPDATE or DELETE.
' - CloseDB:            Close the database.
' - AddWithValue(sParameterName, vValue):       Used to set the parameters in the query type INSERT, UPDATE or DELETE.
' - GetInfo:            Return ALL data from the cSQLiteConnection as a string (optional).
'
'
' Example 1:
'
' `
'  Dim sc As cSQLiteConnection
'  Set sc = New cSQLiteConnection
'
'  sc.OpenDB "c:\database.db"
'  sc.ExecuteQuery "SELECT * FROM MyTable"
'
'  If Not IsEmpty(sc.Cursor) Then
'      Debug.Print sc.CursorString
'  End If
'
'  sc.CloseDB
' `
'
' Example 2:
'
' `
'  Dim sc As cSQLiteConnection
'  Set cs = New cSQLiteConnection
'
'  sc.OpenDB "c:\database.db"
'  sc.Query = "INSERT INTO MyTable ('Arg1','Arg2') VALUES (@Arg1, @Arg2)"
'  sc.AddWithValue("@Arg1", 1)
'  sc.AddWithValue("@Arg2", 2)
'  sc.ExecuteNonQuery
'
'  If Not IsEmpty(sc.Cursor) Then
'      Debug.Print sc.CursorString
'  End If
'
'  sc.CloseDB
' `
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit


' Calling statements for the .dll
Private Declare Sub sqlite3_open Lib "sqlite.dll" (ByVal FileName As String, ByRef Handle As Long)
Private Declare Sub sqlite3_close Lib "sqlite.dll" (ByVal DB_Handle As Long)
Private Declare Function sqlite3_last_insert_rowid Lib "sqlite.dll" (ByVal DB_Handle As Long) As Long
Private Declare Function sqlite3_changes Lib "sqlite.dll" (ByVal DB_Handle As Long) As Long
Private Declare Function sqlite_get_table Lib "sqlite.dll" (ByVal DB_Handle As Long, ByVal SQLString As String, ByRef ErrStr As String) As Variant()
Private Declare Function sqlite_libversion Lib "sqlite.dll" () As String
Private Declare Function number_of_rows_from_last_call Lib "sqlite.dll" () As Long

Private i As Long
Private j As Long
Private k As Long



' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Members
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private m_sFilePath As String
Private m_sQuery As String
Private m_lHandle As Long
Private m_vHeaders As Variant
Private m_vCursor As Variant
Private m_sError As String
Private m_lRows As Long
Private m_lColumns As Long
Private m_lRowLimit As Long
Private m_lColumnLimit As Long
Private m_cItems As Collection

Public Property Get filePath() As String
    filePath = m_sFilePath
End Property

Public Property Let filePath(sValue As String)
    If sValue <> "" And Dir(sValue, vbArchive) <> "" Then
        m_sFilePath = sValue
    End If
End Property


Public Property Get Error() As String
    Error = m_sError
End Property

Private Property Let Error(sValue As String)
    m_sError = sValue
End Property


Public Property Get Cursor() As Variant
    Cursor = m_vCursor
End Property

Public Property Let Cursor(vData As Variant)
    m_vCursor = vData
End Property

Private Property Set Items(cData As Collection)
    Set m_cItems = cData
End Property

Public Property Get Items() As Collection
    Set Items = m_cItems
End Property

Public Property Get CursorString() As String
    Dim sHeaders As String
    
    If Not IsEmpty(Headers) Then
    For i = 0 To UBound(Headers)
        sHeaders = sHeaders + Headers(i) + ","
    Next
    sHeaders = Left(sHeaders, Len(sHeaders) - 1) + vbCrLf
    End If

    
    Dim sRow As String
    Dim sContent As String
    For i = 0 To RowLimit
        sRow = ""
        For j = 0 To ColumnLimit
            sRow = sRow + Cursor(i, j) + ","
        Next
        sRow = Left(sRow, Len(sRow) - 1) + vbCrLf
        sContent = sContent + sRow
    Next
    CursorString = sHeaders + sContent
End Property


Public Property Get Query() As String
    Query = m_sQuery
End Property

Public Property Let Query(sValue As String)
    m_sQuery = sValue
End Property


Public Property Get Headers()
    Headers = m_vHeaders
End Property

Private Property Let Headers(vData As Variant)
    m_vHeaders = vData
End Property


Public Property Get Handle() As Long
    Handle = m_lHandle
End Property

Private Property Let Handle(lValue As Long)
    m_lHandle = lValue
End Property


Public Property Get RowLimit() As Long
    If m_lRowLimit = -1 Then
        RowLimit = m_lRowLimit
    ElseIf IsEmpty(Cursor) Then
        RowLimit = 0
    Else
        RowLimit = UBound(Cursor, 1)
    End If
End Property

Private Property Let RowLimit(lValue As Long)
    m_lRowLimit = lValue
End Property


Public Property Get ColumnLimit() As Long
    If m_lColumnLimit = -1 Then
        ColumnLimit = m_lColumnLimit
    ElseIf IsEmpty(Cursor) Then
        ColumnLimit = 0
    Else
        ColumnLimit = UBound(Cursor, 2)
    End If
End Property

Private Property Let ColumnLimit(lValue As Long)
    m_lColumnLimit = lValue
End Property


Public Property Get Rows() As Long
    Rows = m_lRows
End Property

Private Property Let Rows(lValue As Long)
    m_lRows = lValue
End Property


Public Property Get Columns() As Long
    Columns = m_lColumns
End Property

Private Property Let Columns(lValue As Long)
    m_lColumns = lValue
End Property


Private Sub Class_Initialize()
    Handle = -1
    RowLimit = -1
    ColumnLimit = -1
    Rows = -1
    Columns = -1
    Set Items = New Collection
End Sub



' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub OpenDB(Optional ByVal sFilePath As String = "")
    RowLimit = -1
    ColumnLimit = -1
    Rows = -1
    Columns = -1
    Handle = 0
    filePath = sFilePath
    sqlite3_open m_sFilePath, m_lHandle
End Sub

Public Sub ExecuteQuery(Optional sQuery As Variant = Empty, Optional fIncludeHeaders As Boolean = False)
On Error GoTo CatchError
    If m_lHandle = -1 Then
        Err.Raise vbObjectError + 513, "ExecuteNonQuery", "value of m_lHandle is INVALID. Check that the connection to the database is open"
    End If
    
    Query = sQuery
    m_vCursor = sqlite_get_table(m_lHandle, m_sQuery, m_sError)

On Error GoTo ErrorEmptyCursor

    ColumnLimit = UBound(m_vCursor, 2)
    Columns = ColumnLimit + 1
    Headers = GetHeadersFromCursor
    RowLimit = UBound(m_vCursor, 1)
    
    Set m_cItems = New Collection
    For i = 1 To RowLimit
        Dim oDict As New Dictionary
        
        For j = 0 To ColumnLimit
            oDict.Add m_vCursor(0, j), m_vCursor(i, j)
        Next j

        m_cItems.Add oDict
        Set oDict = Nothing
    Next i
    
    If fIncludeHeaders = False Then
        RemoveHeaderFromCursor
    End If
    
    RowLimit = UBound(Cursor, 1)
    Rows = RowLimit + 1
    
    Exit Sub

ErrorEmptyCursor:
    
    RowLimit = -1
    ColumnLimit = -1
    Rows = -1
    Columns = -1
    Headers = Empty
    Cursor = Empty
    Set Items = Nothing
    
    Exit Sub
    
CatchError:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
End Sub

Public Sub ExecuteNonQuery(Optional fIncludeHeaders As Boolean = False)
On Error GoTo CatchError
    If m_lHandle = -1 Then
        Err.Raise vbObjectError + 513, "ExecuteNonQuery", "value of m_lHandle is INVALID. Check that the connection to the database is open"
    End If

    m_vCursor = sqlite_get_table(m_lHandle, m_sQuery, m_sError)

On Error GoTo ErrorEmptyCursor

    ColumnLimit = UBound(Cursor, 2)
    Columns = ColumnLimit + 1
    Headers = GetHeadersFromCursor
    RowLimit = UBound(m_vCursor, 1)

    Set m_cItems = New Collection
    For i = 1 To RowLimit
        Dim oDict As New Dictionary
        
        For j = 0 To ColumnLimit
            oDict.Add m_vCursor(0, j), m_vCursor(i, j)
        Next j

        m_cItems.Add oDict
        Set oDict = Nothing
    Next i
    
    If fIncludeHeaders = False Then
        RemoveHeaderFromCursor
    End If
    
    RowLimit = UBound(Cursor, 1)
    Rows = RowLimit + 1
    
    Exit Sub
    
ErrorEmptyCursor:
    RowLimit = -1
    ColumnLimit = -1
    Rows = -1
    Columns = -1
    Headers = Empty
    Cursor = Empty
    Set Items = Nothing
    
    Exit Sub
    
CatchError:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
End Sub

Public Sub CloseDB()
    sqlite3_close m_lHandle
    
    filePath = ""
    Handle = -1
    Query = ""
    RowLimit = -1
    ColumnLimit = -1
    Rows = -1
    Columns = -1
    Headers = Empty
    Set Items = Nothing
End Sub

Public Sub AddWithValue(sParameterName As String, vValue As Variant)
    If IsNull(vValue) Then
        Query = Replace(Query, sParameterName, "NULL")
    Else
        Query = Replace(Query, sParameterName, "'" & CStr(vValue) & "'")
    End If
    'Query = Replace(Query, sParameterName, "'" & CStr(vValue) & "'")
End Sub



' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Function GetHeadersFromCursor() As Variant
    Dim vHeaders As Variant
    
    ReDim vHeaders(0 To UBound(Cursor, 2))
    
    For i = 0 To UBound(Cursor, 2)
        vHeaders(i) = Cursor(0, i)
    Next
    
    GetHeadersFromCursor = vHeaders
End Function

Private Function RemoveHeaderFromCursor()
    Dim vNewCursor As Variant
    
    ReDim vNewCursor(0 To UBound(Cursor, 1) - 1, 0 To UBound(Cursor, 2))
    
    For i = 1 To UBound(Cursor, 1)
        For j = 0 To UBound(Cursor, 2)
            vNewCursor(i - 1, j) = Cursor(i, j)
        Next
    Next
    
    Cursor = vNewCursor
End Function



' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Debugging
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function GetInfo(Optional sMessage As String = "") As String
    Dim sMark As String
    If sMessage <> "" Then
        sMark = sMessage
    End If
    Dim sHeaders As String
    If IsEmpty(Headers) = False Then
        For i = 0 To UBound(Cursor, 2)
            sHeaders = sHeaders + Headers(i) + ","
        Next
        sHeaders = Left(sHeaders, Len(sHeaders) - 1)
    Else
        sHeaders = "Empty"
    End If
    
    GetInfo = "" _
    & "+----------- " & CStr(Time) & " ------------" & sMark & vbCrLf _
    & "| FilePath: " & """" & filePath & """" & vbCrLf _
    & "| Handle: " & Handle & vbCrLf _
    & "| Query: " & """" & Query & """" & vbCrLf _
    & "| RowLimit: " & RowLimit & vbCrLf _
    & "| ColumnLimit: " & ColumnLimit & vbCrLf _
    & "| Rows: " & Rows & vbCrLf _
    & "| Columns: " & Columns & vbCrLf _
    & "| Headers: " & sHeaders & vbCrLf
End Function
