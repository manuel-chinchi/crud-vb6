VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' file:     clsToolTip.cls
' author:   Manuel Chinchi
' date:     30/11/2024
'
' Esta clase representa un tooltip personalizado tipo burbuja que
' aparece debajo del control padre.
'******************************************************************************

Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
    ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, _
    ByVal hInstance As Long, ByVal lpParam As Any) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Any) As Long

Private Declare Function GetClientRect Lib "user32.dll" ( _
    ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function DestroyWindow Lib "user32" ( _
    ByVal hwnd As Long) As Long

'Private Declare Function InvalidateRect Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    lpRect As Any, _
    ByVal bErase As Long) As Long

Private Const TTS_NOPREFIX As Long = &H2
Private Const WS_POPUP = &H80000000
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_BALLOON = &H40
Private Const TTS_TRACK As Long = &H20 ' Para habilitar el seguimiento manual
Private Const TTM_ADDTOOL = &H432
Private Const TTM_SETMAXTIPWIDTH = &H418
Private Const TTM_TRACKPOSITION As Long = &H412
Private Const TTM_TRACKACTIVATE As Long = &H411
Private Const TTF_TRANSPARENT As Long = &H100
Private Const TTF_CENTERTIP As Long = &H2

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

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private hTooltip As Long
Public mLabel As Label
Public mParent As Object
Public mTextBox As TextBox

Public Sub ShowToolTip_OLD(sText As String)

    hTooltip = CreateWindowEx(0, "tooltips_class32", vbNullString, WS_POPUP Or TTS_ALWAYSTIP Or TTS_BALLOON, _
                              0, 0, 0, 0, mParent.hwnd, 0, App.hInstance, ByVal 0&)
    Dim ti As TOOLINFO
    Dim rc As RECT
    Dim maxWidth As Long
    maxWidth = 300

    ' Define las dimensiones del Label
    'rc.Left = mLabel.Left
    'rc.Top = mLabel.Top
    'rc.Right = mLabel.Left + mLabel.Width
    'rc.Bottom = mLabel.Top + mLabel.Height
    
    rc.Left = mLabel.Left / Screen.TwipsPerPixelX
    rc.Top = mLabel.Top / Screen.TwipsPerPixelY
    rc.Right = (mLabel.Left + mLabel.Width) / Screen.TwipsPerPixelX
    rc.Bottom = (mLabel.Top + mLabel.Height) / Screen.TwipsPerPixelY

    ' Configura el Tooltip
    With ti
        .cbSize = Len(ti)
        .hwnd = mParent.hwnd
        .RECT = rc
        '.lpszText = "Esto es un Tooltip avanzado"
        .lpszText = sText
    End With

    'SendMessage hTooltip, TTM_ADDTOOL, 0, TOOLINFO
    SendMessage hTooltip, TTM_ADDTOOL, 0, VarPtr(ti)
    'SendMessage hTooltip, TTM_SETMAXTIPWIDTH, 0, ByVal 300
    SendMessage hTooltip, TTM_SETMAXTIPWIDTH, 0, VarPtr(maxWidth)
    
    SendMessage hTooltip, TTM_TRACKPOSITION, 0, ByVal (rc.Left + (rc.Right - rc.Left) \ 2) Or (rc.Top + (rc.Bottom - rc.Top) \ 2) * &H10000
    SendMessage hTooltip, TTM_TRACKACTIVATE, 1, VarPtr(ti)
    'Debug.Print Now
End Sub

Public Sub ShowToolTip(sText As String)
    ' Crea el Tooltip
    hTooltip = CreateWindowEx(0, "tooltips_class32", vbNullString, WS_POPUP Or TTS_ALWAYSTIP Or TTS_BALLOON, _
                              0, 0, 0, 0, mTextBox.hwnd, 0, App.hInstance, ByVal 0&)
    Dim ti As TOOLINFO
    Dim rc As RECT
    Dim point As POINTAPI
    Dim maxWidth As Long
    maxWidth = 300

    ' Define las dimensiones del Label o TextBox
    point.x = mTextBox.Left / Screen.TwipsPerPixelX
    point.y = mTextBox.Top / Screen.TwipsPerPixelY

    ' Convierte las coordenadas del cliente a coordenadas de pantalla
    Call ClientToScreen(mTextBox.hwnd, point)

    ' Configura el Tooltip
    With ti
        .cbSize = Len(ti)
        .hwnd = mTextBox.hwnd
        .RECT.Left = 0
        .RECT.Top = 0
        .RECT.Right = 0
        .RECT.Bottom = 0
        .lpszText = sText
    End With

    ' Asocia el tooltip con el control
    SendMessage hTooltip, TTM_ADDTOOL, 0, VarPtr(ti)
    SendMessage hTooltip, TTM_SETMAXTIPWIDTH, 0, ByVal maxWidth

    ' Muestra el tooltip en la esquina superior izquierda del control
    'SendMessage hTooltip, TTM_TRACKPOSITION, 0, ByVal (point.X Or (point.Y * &H10000))
    SendMessage hTooltip, TTM_TRACKACTIVATE, 1, VarPtr(ti)
End Sub

Public Sub ShowToolTipFixed(sText As String, ByVal x As Long, ByVal y As Long, Optional hwnd As Long)
    ' Crea el Tooltip
    If hTooltip <> 0 Then
        DestroyWindow hTooltip
        hTooltip = 0
        Debug.Print "tooltip destruido"
    End If
    
    hTooltip = CreateWindowEx( _
        0, _
        "tooltips_class32", _
        vbNullString, _
        WS_POPUP Or TTS_ALWAYSTIP Or TTS_BALLOON Or TTS_TRACK, _
        0, 0, 0, 0, _
        mParent.hwnd, _
        0, _
        App.hInstance, _
        ByVal 0&)
    'WS_POPUP Or TTS_ALWAYSTIP Or TTS_BALLOON Or TTS_TRACK, _

    Dim ti As TOOLINFO
    Dim maxWidth As Long
    maxWidth = 300

    ' Configura el Tooltip
    With ti
        .cbSize = Len(ti)
        .hwnd = mParent.hwnd
        .uFlags = TTF_TRANSPARENT Or TTF_CENTERTIP
        .RECT.Left = 0
        .RECT.Top = 0
        .RECT.Right = 0
        .RECT.Bottom = 0
        .lpszText = sText & " " & Str(hTooltip)
    End With
    
    GetClientRect mParent.hwnd, ti.RECT
    
    ' Asocia el tooltip
    SendMessage hTooltip, TTM_ADDTOOL, 0, VarPtr(ti)
    'SendMessage hTooltip, TTM_SETMAXTIPWIDTH, 0, ByVal maxWidth
'x = 0
'y = 0
    ' Establece la posición fija
    SendMessage hTooltip, TTM_TRACKPOSITION, 0, ByVal (x Or (y * &H10000))
    
    ' Activa el tooltip en modo track
    SendMessage hTooltip, TTM_TRACKACTIVATE, 1, VarPtr(ti)
End Sub

Public Sub HideToolTip()
    Debug.Print "hidetooltip: " & hTooltip
    
    ' Destruir el tooltip activo, si existe
    If hTooltip <> 0 Then
        SendMessage hTooltip, TTM_TRACKACTIVATE, 0, ByVal 0&
        
        DestroyWindow hTooltip
        hTooltip = 0
        
        'Dim rc As RECT
        'GetClientRect mParent.hWnd, rc ' O usa hwnd del control si es específico
        'InvalidateRect mParent.hWnd, rc, True
    End If
    
   
End Sub
