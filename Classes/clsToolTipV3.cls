VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTipV3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
    ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, _
    ByVal hInstance As Long, ByVal lpParam As Any) As Long

Private Declare Function GetClientRect Lib "user32.dll" ( _
    ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Any) As Long
    
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTS_BALLOON As Long = &H40
Private Const TTF_TRANSPARENT As Long = &H100
Private Const TTF_CENTERTIP As Long = &H2
Private Const WS_POPUP As Long = &H80000000
Private Const TOOLTIPS_CLASS As String = "tooltips_class32"
Private Const TTM_TRACKPOSITION As Long = &H412
Private Const TTM_TRACKACTIVATE As Long = &H411
Private Const TTM_SETMAXTIPWIDTH = &H418
Private Const TTM_ADDTOOL As Long = &H400 + 4
Private Const LPSTR_TEXTCALLBACK As Long = -1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uId As Long
    RECT As RECT
    hinst As Long
    lpszText As String
End Type

Public mParent As Object
Public mControl As Object

Public Sub ShowToolTip(sText As String, hwnd As Long)
    Dim hwndTooltip As Long
    Dim ti As TOOLINFO
    Dim rc As RECT

    ' Crear la ventana Tooltip
    hwndTooltip = CreateWindowEx( _
        0, _
        TOOLTIPS_CLASS, _
        vbNullString, _
        WS_POPUP Or TTS_NOPREFIX Or TTS_BALLOON, _
        0, 0, 0, 0, _
        hwnd, _
        0, _
        App.hInstance, _
        ByVal 0&)
    
    If hwndTooltip Then
        ' Configurar TOOLINFO
        With ti
            .cbSize = Len(ti)
            .uFlags = TTF_TRANSPARENT Or TTF_CENTERTIP
            '.uFlags = TTF_TRANSPARENT
            .hwnd = hwnd
            .uId = 0
            .hinst = 0
            .lpszText = "ABC" ' Para callback
        End With

        ' Obtener el área del control (adaptar según sea necesario)
        GetClientRect hwnd, rc
        ti.RECT = rc
        ti.RECT.Right = 5
        ti.RECT.Left = 5

        ' Asociar la herramienta al Tooltip
        SendMessage hwndTooltip, TTM_ADDTOOL, 0, ByVal VarPtr(ti)
        
        ' ++++++
        SendMessage hwndTooltip, TTM_TRACKACTIVATE, 1, VarPtr(ti)
    End If
End Sub









