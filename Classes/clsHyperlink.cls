VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const IDC_HAND = 32649&
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
     ByVal hwnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long

Private WithEvents mLabel As Label
Attribute mLabel.VB_VarHelpID = -1
Public URL As String

Private mBrowsers As New Collection

Const PATH_IE_X32 = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
Const PATH_IE_X64 = "C:\Program Files\Internet Explorer\iexplore.exe"
Const PATH_FIREFOX_X32 = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
Const PATH_FIREFORX_X64 = "C:\Program Files\Mozilla Firefox\firefox.exe"
Const PATH_CHROME_X32 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
Const PATH_CHROME_X64 = "C:\Program Files\Google\Chrome\Application\chrome.exe"

Public Sub Initialize(ByRef lbl As Object)
    Set mLabel = lbl

    With mLabel
        .FontUnderline = True
        .ForeColor = &H8000000D
    End With

    If Dir(PATH_FIREFOX_X32) <> "" Then
        mBrowsers.Add PATH_FIREFOX_X32
    End If
    If Dir(PATH_FIREFORX_X64) <> "" Then
        mBrowsers.Add PATH_FIREFORX_X64
    End If

    If Dir(PATH_CHROME_X32) <> "" Then
        mBrowsers.Add PATH_CHROME_X32
    End If
    If Dir(PATH_CHROME_X64) <> "" Then
        mBrowsers.Add PATH_CHROME_X64
    End If
    
    If Dir(PATH_IE_X32) <> "" Then
        mBrowsers.Add PATH_IE_X32
    End If
    If Dir(PATH_IE_X64) <> "" Then
        mBrowsers.Add PATH_IE_X64
    End If
End Sub

Private Sub mLabel_Click()
    If mBrowsers.Count > 0 Then
        ShellExecute 0, "open", mBrowsers.Item(1), Me.URL, vbNullString, 1
    Else
        MsgBox "View the URL " & Me.URL, vbOKOnly
    End If
End Sub

Private Sub mLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Public Sub ShowInstalledBrowsers()
    Dim i
    For i = 1 To mBrowsers.Count
        Debug.Print i & " : " & mBrowsers.Item(i)
    Next
End Sub

