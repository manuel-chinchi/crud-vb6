VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main"
   ClientHeight    =   3372
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3624
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3372
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
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   732
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

Private mFormUIManager As New clsFormUIManager
Private mHyperlink As New clsHyperlink

Private Sub cmdArticles_Click()
    frmListArticles.Show vbModal
End Sub

Private Sub cmdCategories_Click()
    frmListCategories.Show vbModal
End Sub

Private Sub cmdReports_Click()
    frmReports.Show vbModal
End Sub

Private Sub Form_Initialize()
    mFormUIManager.Initialize Me
    mHyperlink.Initialize Me.lblAbout
    mHyperlink.URL = "https://github.com/manuel-chinchi/crud-vb6"
End Sub

