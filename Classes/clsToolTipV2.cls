VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTipV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ECM_FIRST = &H1500                    '// Edit control messages

Private Const EM_SETCUEBANNER = (ECM_FIRST + 1)
Private Const EM_GETCUEBANNER = (ECM_FIRST + 2)     '// Set the cue banner with the lParm = LPCWSTR

Private Type EDITBALLOONTIP
    cbStruct As Long
    pszTitle As Long
    pszText As Long
    ttiIcon As Long ' ; // From TTI_*
End Type

Private Const EM_SHOWBALLOONTIP = (ECM_FIRST + 3)      '// Show a balloon tip associated to the edit control
Private Const EM_HIDEBALLOONTIP = (ECM_FIRST + 4)       '// Hide any balloon tip associated with the edit control
Private Declare Function SendMessageW Lib "user32" ( _
   ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private m_hWnd As Long
Private m_sCueBanner As String
Private m_sTitle As String
Private m_sText As String
Private m_eIcon As String

Public Property Let TextBox(txtThis As TextBox)
   m_hWnd = txtThis.hwnd
End Property

Public Property Let CueBanner(ByVal value As String)
   m_sCueBanner = value
   setCueBanner
End Property
Public Property Get CueBanner() As String
   CueBanner = m_sCueBanner
End Property
Public Property Let BalloonTipTitle(ByVal value As String)
   m_sTitle = value
End Property
Public Property Get BalloonTipTitle() As String
   BalloonTipTitle = m_sTitle
End Property
Public Property Let BalloonTipText(ByVal value As String)
   m_sText = value
End Property
Public Property Get BalloonTipText() As String
   BalloonTipText = m_sText
End Property
Public Property Let BalloonTipIcon(ByVal value As String)
   m_eIcon = value
End Property
Public Property Get BalloonTipIcon() As String
   BalloonTipIcon = m_eIcon
End Property

Public Sub ShowBalloonTip()

    Dim lR As Long
   Dim tEBT As EDITBALLOONTIP
   tEBT.cbStruct = LenB(tEBT)
   tEBT.pszText = StrPtr(m_sText)
   tEBT.pszTitle = StrPtr(m_sTitle)
   'tEBT.ttiIcon = m_eIcon
   lR = SendMessageW(m_hWnd, EM_SHOWBALLOONTIP, 0, tEBT)
End Sub
Public Sub HideBalloonTip()
    Dim lR As Long
   lR = SendMessageLongW(m_hWnd, EM_HIDEBALLOONTIP, 0, 0)
   Debug.Print lR
End Sub
Private Sub setCueBanner()
Dim lR As Long
   ' Reports success, but doesn'/t actually work...
   ' (is this because the VB text box is ANSI?)
   lR = SendMessageLongW(m_hWnd, EM_SETCUEBANNER, 0, StrPtr(m_sCueBanner))
   Debug.Print lR
End Sub

