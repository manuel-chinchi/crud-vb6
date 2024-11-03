VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDependencyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents mMainForm As Form
Attribute mMainForm.VB_VarHelpID = -1

Public Sub Initialize(ByRef frm As Object)
    Set mMainForm = frm
End Sub

Private Sub mMainForm_Load()
    FileCopy App.Path & "\Dependences\SQLite\sqlite.dll", App.Path & "\sqlite.dll"
End Sub

Private Sub mMainForm_Unload(Cancel As Integer)
   'TODO El archivo 'sqlite.dll' no se puede borrar desde el programa principal
   ' ya que el modulo cSQLiteConnection lo usa y por ende lo bloquea hasta que
   ' se cierra. Esta es una forma rápidad de solucionar eso. Revisar a futuro.
   
    Shell App.Path & "\Scripts\cleanup.bat", vbNormal
End Sub
