VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Dim WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1

Public Sub Initialize(ByRef frm As Object)
    Set mForm = frm
End Sub

Private Sub mForm_Load()
    InitCommonControls
End Sub
