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
Dim mPathSqliteDependency As String
Dim mPathCleanFilesScript As String

Public Sub Initialize(ByRef frm As Object)
    Set mMainForm = frm
    
    mPathSqliteDependency = App.Path & "\Dependences\SQLite\sqlite.dll"
    mPathCleanFilesScript = App.Path & "\Scripts\cleanup.bat"
End Sub

Private Sub mMainForm_Load()
    LoadDependences
End Sub

Private Sub mMainForm_Unload(Cancel As Integer)
    CleanTempFiles
End Sub

Private Sub LoadDependences()
On Error GoTo FileNotFound
    FileCopy mPathSqliteDependency, App.Path & "\sqlite.dll"
    
    Exit Sub
    
FileNotFound:
    Err.Raise vbObjectError + 513, "LoadDependences", "The file " & mPathSqliteDependency & " was not found"
End Sub

Private Sub CleanTempFiles()
    'TODO El archivo 'sqlite.dll' no se puede borrar desde el programa principal
   ' ya que el modulo cSQLiteConnection lo usa y por ende lo bloquea hasta que
   ' se cierra. Esta es una forma rápidad de solucionar eso. Revisar a futuro.

On Error GoTo FileNotFound
    Shell mPathCleanFilesScript, vbNormal
    
    Exit Sub

FileNotFound:
    Err.Raise vbObjectError + 513, "CleanTempFiles", "The file " & mPathCleanFilesScript & " was not found"
End Sub
