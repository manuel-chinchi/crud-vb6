VERSION 5.00
Begin VB.Form frmCreateArticle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CreateArticle"
   ClientHeight    =   4788
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3624
   Icon            =   "frmCreateArticle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4788
   ScaleWidth      =   3624
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboCategories 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   120
      TabIndex        =   2
      Text            =   "---Seleccionar---"
      Top             =   2640
      Width           =   3372
   End
   Begin VB.TextBox txtDetails 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3372
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3372
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3372
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   3372
   End
   Begin VB.Label lblCategory 
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1092
   End
   Begin VB.Label lblDetails 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "ARTICLE DATA"
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
      TabIndex        =   4
      Top             =   240
      Width           =   3612
   End
End
Attribute VB_Name = "frmCreateArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mArticle As New clsArticle
Dim mDialogResult As VbMsgBoxResult
Dim mCategoryRepository As clsCategoryRepository
Dim mComboBoxUIManager As New clsComboBoxUIManager
Private tt As clsToolTip
Private ttName As New clsToolTip
Private ttDetails As New clsToolTip

Public Property Get Article() As clsArticle
    Set Article = mArticle
End Property

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = mDialogResult
End Property

Public Property Get CategoryId() As Integer
    Dim oCategory As clsCategory
    Dim cCategories As New Collection
    
    Set cCategories = mCategoryRepository.GetCategories()
    For Each oCategory In cCategories
        If cboCategories.Text = oCategory.mName Then
            CategoryId = oCategory.mId
            Exit For
        End If
    Next oCategory
End Property

Private Sub Form_Load()
On Error GoTo Catch
    Set mCategoryRepository = modSingletonRepository.GetCategoryRepository()

    If mCategoryRepository.GetCategories().Count <> 0 Then
        Dim vCategories() As Variant
        
        vCategories = modCategoryHelper.ConvertToVariant(mCategoryRepository.GetCategories())
        SetComboBox vCategories
    Else
        cboCategories.Enabled = False
    End If
    
    cboCategories.ListIndex = 0
    
    mComboBoxUIManager.Initialize cboCategories
    Exit Sub
    
Catch:
     
     Err.Raise vbObjectError + 513, , Err.Number & " " & Err.Description & " on Form_Load (" & Me.Name & ")"
End Sub

Private Sub cmdAccept_Click()
    With mArticle
        .mId = 0
        .mName = txtName.Text
        .mDetails = txtDetails.Text
        If .mCategory Is Nothing Then
            Set .mCategory = New clsCategory
        End If
        .mCategory.mName = cboCategories.Text
        .mCategory.mId = Me.CategoryId
    End With
    
    mDialogResult = vbOK
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mDialogResult = vbCancel
    
    Unload Me
End Sub

Private Sub SetComboBox(ParamArray varParam() As Variant)
    Dim i
    Dim v As Variant
    
    For Each v In varParam
        If IsArray(v) Then
            For i = LBound(v) To UBound(v)
                cboCategories.AddItem v(i)
            Next
        Else
            cboCategories.AddItem v
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case vbFormControlMenu 'X'
            mDialogResult = vbCancel
    End Select
End Sub

' ***
Private Sub txtName_KeyPress(KeyAscii As Integer)
    'If Not tt Is Nothing Then
    '    tt.HideToolTip
    'End If
    ttName.HideToolTip
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    If txtName.Text = "" Then
        Set ttName.mParent = txtName
        ttName.ShowToolTipFixed "111", 0, 0
        Cancel = True
    Else
        ttName.HideToolTip
    End If
    
    Exit Sub
    Set tt = New clsToolTip
    If txtName.Text = "" Then
        Set tt.mParent = txtName
        tt.ShowToolTipFixed "El campo nombre no puede estar vacío", 0, 0
        Cancel = True
    Else
        tt.HideToolTip
    End If
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
    'If Not tt Is Nothing Then
    '    tt.HideToolTip
    'End If
    ttDetails.HideToolTip
End Sub

Private Sub txtDetails_Validate(Cancel As Boolean)
    If txtDetails.Text = "" Then
        Set ttDetails.mParent = txtDetails
        ttDetails.ShowToolTipFixed "xxx", 0, 0
        Cancel = True
    Else
        ttDetails.HideToolTip
    End If
    
    Exit Sub
    If txtDetails.Text = "" Then
        Set tt.mParent = txtDetails
        tt.ShowToolTipFixed "El campo Detalles no puede estar vacío", 0, 0
        Cancel = True
    Else
        tt.HideToolTip
    End If
End Sub

