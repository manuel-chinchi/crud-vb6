VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListCategories 
   Caption         =   "ListCategories"
   ClientHeight    =   4356
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8184
   LinkTopic       =   "Form1"
   ScaleHeight     =   4356
   ScaleWidth      =   8184
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwCategories 
      Height          =   3612
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   6371
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6120
      TabIndex        =   1
      Top             =   1320
      Width           =   1932
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6120
      TabIndex        =   0
      Top             =   600
      Width           =   1932
   End
End
Attribute VB_Name = "frmListCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
