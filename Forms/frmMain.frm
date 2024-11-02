VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main"
   ClientHeight    =   3144
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3144
   ScaleWidth      =   3624
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReports 
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   3612
   End
   Begin VB.CommandButton cmdCategories 
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   3612
   End
   Begin VB.CommandButton cmdArticles 
      Caption         =   "Articles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3612
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   3612
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Sub cmdArticles_Click()
    frmListArticles.Show vbModal
End Sub

Private Sub cmdCategories_Click()
    frmListCategories.Show vbModal
End Sub

Private Sub cmdReports_Click()
    frmReports.Show vbModal
End Sub

Private Sub Form_Load()
    InitializeRepositories
    
    'FileCopy App.Path & "\Dependences\SQLite\sqlite.dll", App.Path & "\sqlite.dll"
    CopyFile App.Path & "\Dependences\SQLite\sqlite.dll", App.Path & "\sqlite.dll", False
End Sub

Private Sub InitializeRepositories()
    Set frmCreateArticle.CategoryRepository = modSingletonRepository.GetCategoryRepository()
    Set frmListArticles.ArticleRepository = modSingletonRepository.GetArticleRepository()
    Set frmListCategories.CategoryRepository = modSingletonRepository.GetCategoryRepository()
    Set frmEditArticle.CategoryRepository = modSingletonRepository.GetCategoryRepository()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteAFile App.Path & "\sqlite.dll"
End Sub

Sub DeleteAFile(filePath)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   SetAttr filePath, vbNormal
   fso.DeleteFile (filePath)
End Sub
