Attribute VB_Name = "modIOHelper"
Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const INVALID_HANDLE_VALUE As Long = -1

Dim iFile

Function IsDllInUse(ByVal filePath As String) As Boolean
    Dim hFile As Long

    ' Intentar abrir el archivo en modo exclusivo
    hFile = CreateFile(filePath, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        ' No se pudo abrir el archivo, está en uso
        IsDllInUse = True
    Else
        ' Se pudo abrir el archivo, no está en uso
        IsDllInUse = False
        CloseHandle hFile
    End If
End Function

Function FreeFileEx(ByVal hLib As Long) As Long
    FreeFileEx = FreeLibrary(hLib)
End Function


Public Function GetStringFromFile(strFilename As String) As String
  iFile = FreeFile
  Open strFilename For Input As #iFile
    GetStringFromFile = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function
