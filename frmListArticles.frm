VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListArticles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ListArticles"
   ClientHeight    =   5772
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   9408
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   9408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   5892
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1092
   End
   Begin MSComctlLib.ListView lvwArticles 
      Height          =   5052
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7092
      _ExtentX        =   12510
      _ExtentY        =   8911
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "Show All"
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
      Left            =   7320
      TabIndex        =   3
      Top             =   2760
      Width           =   1932
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
      Left            =   7320
      TabIndex        =   2
      Top             =   2040
      Width           =   1932
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      Left            =   7320
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
      Left            =   7320
      TabIndex        =   0
      Top             =   600
      Width           =   1932
   End
End
Attribute VB_Name = "frmListArticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Articles As Collection
Dim i As Integer

Private Sub Form_Load()
    Set Articles = New Collection
    
    Articles.Add modArticleHelper.NewArticle(1, "buzo", "5xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(2, "remera", "20xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(300, "jean", "40xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(4, "medias", "400xU", "calzado")
    
    SetHeader "Id", "Name", "Details", "Category"
    SetHeaderWidth 900, 1500, 1800, 1200
    SetDataSource Articles
End Sub

Private Sub cmdAdd_Click()
    frmCreateArticle.Show vbModal
    Articles.Add frmCreateArticle.Article
    SetDataSource Articles
End Sub

Private Sub cmdEdit_Click()
    Dim oArticle As clsArticle
    Set oArticle = New clsArticle
    Dim li As ListItem
    Set li = lvwArticles.SelectedItem
        
    With oArticle
        .mId = li.Text
        .mName = li.SubItems(1)
        .mDetails = li.SubItems(2)
        .mCategoryName = li.SubItems(3)
    End With
    
    Set frmEditArticle.Article = oArticle
    frmEditArticle.Show vbModal
    
    If frmEditArticle.DialogResult = vbOK Then
        Dim oCurrentArticle As clsArticle
        For Each oCurrentArticle In Articles
            If oCurrentArticle.mId = frmEditArticle.Article.mId Then
                With oCurrentArticle
                    .mName = frmEditArticle.Article.mName
                    .mDetails = frmEditArticle.Article.mDetails
                    .mCategoryName = frmEditArticle.Article.mCategoryName
                End With
                Exit For
            End If
        Next
    End If
    
    SetDataSource Articles
End Sub

Private Sub cmdDelete_Click()
    If Not lvwArticles.SelectedItem Is Nothing Then
        Articles.Remove Int(lvwArticles.SelectedItem.Index)
        SetDataSource Articles
    End If
End Sub

Private Sub SetHeader(ParamArray varParam() As Variant)
    With lvwArticles
        With .ColumnHeaders
            .Clear
            
            For i = 0 To UBound(varParam)
                .Add , , varParam(i), 1000
            Next
            
        End With
    End With
End Sub

Private Sub SetHeaderWidth(ParamArray varParam() As Variant)
    With lvwArticles
        With .ColumnHeaders
            
            For i = 0 To UBound(varParam)
                .Item(i + 1).Width = varParam(i)
            Next
            
        End With
    End With
End Sub

Private Sub SetDataSource(arr As Collection)
    Dim li As ListItem
    Dim objArticle As clsArticle

    lvwArticles.ListItems.Clear
    
    For Each objArticle In arr
        Set li = lvwArticles.ListItems.Add(, , objArticle.mId)
        li.SubItems(1) = objArticle.mName
        li.SubItems(2) = objArticle.mDetails
        li.SubItems(3) = objArticle.mCategoryName
    Next
End Sub
