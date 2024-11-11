VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboBoxUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents mComboBox As ComboBox
Attribute mComboBox.VB_VarHelpID = -1

Public Sub Initialize(ByRef cbo As Object)
    Set mComboBox = cbo
End Sub

Private Sub mComboBox_KeyUp(KeyCode As Integer, Shift As Integer)
    AutoSel mComboBox, KeyCode
End Sub

' Thanks @Danial, original code: https://www.vbforums.com/showthread.php?303055-VB6-Auto-Complete-Combo-Box
Private Sub AutoSel(Cmb As ComboBox, KeyCode As Integer)
    If KeyCode = vbEnter Then Exit Sub
    If KeyCode = 8 Then Exit Sub        'Backspace
    If KeyCode = 37 Then Exit Sub       'left key
    If KeyCode = 38 Then Exit Sub       'up arrow key
    If KeyCode = 39 Then Exit Sub       'right key
    If KeyCode = 40 Then Exit Sub       'down arrow key
    If KeyCode = 46 Then Exit Sub       'delete key
    If KeyCode = 33 Then Exit Sub       'page up key
    If KeyCode = 34 Then Exit Sub       'page down key
    If KeyCode = 35 Then Exit Sub       'end key
    If KeyCode = 36 Then Exit Sub       'home key
    
    
    Dim Text As String
    Text = Cmb.Text
    
    Dim i As Long
    Dim Temp As String
    
    
    For i = 0 To Cmb.ListCount
        Temp = Left(Cmb.List(i), Len(Text))
        If LCase(Temp) = LCase(Text) Then
            Cmb.Text = Cmb.List(i)
            Cmb.ListIndex = i
            Cmb.SelStart = Len(Text)
            Cmb.SelLength = Len(Cmb.List(i))
            'Cmb.SetFocus
        End If
    Next
End Sub
