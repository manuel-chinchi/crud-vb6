VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListCategories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ListCategories"
   ClientHeight    =   4356
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   8184
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4356
   ScaleWidth      =   8184
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwCategories 
      Height          =   3612
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   6371
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
Option Explicit

Dim Categories As Collection
Dim i As Integer

Private Sub Form_Load()
    Set Categories = New Collection
    
    Dim Articles As Collection
    Set Articles = New Collection
    Articles.Add modArticleHelper.NewArticle(1, "buzo", "5xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(2, "remera", "20xU", "indumentaria")
    Articles.Add modArticleHelper.NewArticle(300, "jean", "40xU", "indumentaria")
    
    Categories.Add modCategoryHelper.NewCategory(1, "indumentaria", Articles)
    Categories.Add modCategoryHelper.NewCategory(2, "remeras", Nothing)
    Categories.Add modCategoryHelper.NewCategory(3, "pantalones", Nothing)
    
    Set Articles = New Collection
    Articles.Add modArticleHelper.NewArticle(4, "medias", "400xU", "calzado")
    
    Categories.Add modCategoryHelper.NewCategory(4, "calzado", Articles)
    
    SetHeader "Id", "Name", "Articles"
    SetHeaderWidth 900, 1800, 1200
    SetDataSource Categories
End Sub

Private Sub cmdAdd_Click()
    frmCreateCategory.Show vbModal
    
    If frmCreateCategory.DialogResult = vbOK Then
        Categories.Add frmCreateCategory.Category
        SetDataSource Categories
    End If
End Sub

Private Sub cmdDelete_Click()
    If Not lvwCategories.SelectedItem Is Nothing Then
        Categories.Remove (lvwCategories.SelectedItem.Index)
        SetDataSource Categories
    End If
End Sub

Private Sub SetHeader(ParamArray varParam() As Variant)
    With lvwCategories
        With .ColumnHeaders
            .Clear
            
            For i = 0 To UBound(varParam)
                .Add , , varParam(i), 1000
            Next
            
        End With
    End With
End Sub

Private Sub SetHeaderWidth(ParamArray varParam() As Variant)
    With lvwCategories
        With .ColumnHeaders
            For i = 0 To UBound(varParam)
                .Item(i + 1).Width = varParam(i)
            Next
        End With
    End With
End Sub

Private Sub SetDataSource(arr As Collection)
    Dim li As ListItem
    Dim objCategory As clsCategory

    lvwCategories.ListItems.Clear
    
    For Each objCategory In arr
        Set li = lvwCategories.ListItems.Add(, , objCategory.mId)
        li.SubItems(1) = objCategory.mName
        If Not objCategory.mArticlesRelated Is Nothing Then
            li.SubItems(2) = objCategory.mArticlesRelated.Count
        Else
            li.SubItems(2) = 0
        End If
    Next
End Sub
