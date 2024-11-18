VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListCategories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ListCategories"
   ClientHeight    =   4356
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   8184
   Icon            =   "frmListCategories.frx":0000
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
      Checkboxes      =   -1  'True
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

Dim i As Integer
Dim mCategoryRepository As clsCategoryRepository
Dim mListViewUIManager As New clsListViewUIManager

Private Sub Form_Load()
    Set mCategoryRepository = modSingletonRepository.GetCategoryRepository()
    
    LoadCategoryHeaders
    LoadCategories mCategoryRepository.GetCategories()
    cmdDelete.Enabled = False
    
    mListViewUIManager.Initialize lvwCategories
End Sub

Private Sub cmdAdd_Click()
    frmCreateCategory.Show vbModal
    
    If frmCreateCategory.DialogResult = vbOK Then
        mCategoryRepository.CreateCategory frmCreateCategory.Category
        LoadCategories mCategoryRepository.GetCategories()
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim arrIdsCategoriesSelected As Collection
    Set arrIdsCategoriesSelected = GetIdsOfSelectedCategories
    Dim iId As Variant
    
    If arrIdsCategoriesSelected.Count <> 0 Then
        Dim AnswerResult As VbMsgBoxResult
        AnswerResult = MsgBox("Do you want to delete the selected items?", vbExclamation + vbYesNo, "Delete Article")
        If AnswerResult = vbNo Then Exit Sub
    
        For Each iId In arrIdsCategoriesSelected
            mCategoryRepository.DeleteCategory (Int(iId))
        Next
            
        LoadCategories mCategoryRepository.GetCategories()
    End If
End Sub

Private Sub lvwCategories_Click()
    Dim cArray As Collection
    Set cArray = GetIdsOfSelectedCategories
    
    If cArray.Count = 0 Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
End Sub

Private Sub LoadCategoryHeaders()
    With lvwCategories
        With .ColumnHeaders
            .Clear

            .Add , , " ", 300
            .Add , , "Id", 900
            .Add , , "Name", 1800
            .Add , , "Articles", 1200
        End With
    End With
End Sub

Private Sub LoadCategories(arr As Collection)
    Dim li As ListItem
    Dim oCategory As clsCategory

    lvwCategories.ListItems.Clear
    
    For Each oCategory In arr
        Set li = lvwCategories.ListItems.Add(, , "")
        li.SubItems(1) = oCategory.mId
        li.SubItems(2) = oCategory.mName
        If Not oCategory.mArticles Is Nothing Then
            li.SubItems(3) = oCategory.mArticles.Count
        Else
            li.SubItems(3) = 0
        End If
    Next
End Sub

Private Function GetIdsOfSelectedCategories() As Collection
    Dim arrIdsCategories As Collection
    Set arrIdsCategories = New Collection
    Dim li As ListItem
    
    For Each li In lvwCategories.ListItems
        If li.Checked Then
            arrIdsCategories.Add li.SubItems(1)
        End If
    Next
    
    Set GetIdsOfSelectedCategories = arrIdsCategories
End Function
