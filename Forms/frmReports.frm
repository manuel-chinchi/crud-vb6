VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports"
   ClientHeight    =   11292
   ClientLeft      =   36
   ClientTop       =   660
   ClientWidth     =   9852
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11292
   ScaleWidth      =   9852
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgSaveAs 
      Left            =   120
      Top             =   10800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboReports 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   5040
      TabIndex        =   2
      Text            =   "---Seleccionar---"
      Top             =   10900
      Width           =   4692
   End
   Begin CRVIEWERLibCtl.CRViewer crViewer 
      Height          =   10812
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9852
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Label lblChooseReport 
      Alignment       =   1  'Right Justify
      Caption         =   "Choose report:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3480
      TabIndex        =   1
      Top             =   10920
      Width           =   1452
   End
   Begin VB.Menu mnuSaveAs 
      Caption         =   "Save as"
      Begin VB.Menu miSaveAs_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu miSaveAs_PDF 
         Caption         =   "PDF"
      End
      Begin VB.Menu miSaveAs_Word 
         Caption         =   "Word"
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mReport As CRAXDRT.Report
Dim i

Const ZOOM_FULL_WIDTH As Integer = 1
Const ZOOM_FULL_PAGE As Integer = 2

Enum eReport
    rArticleReport = 0
    rCategoriesReport = 1
End Enum

Enum eFormatType
    ftExcel = 0
    ftPDF = 1
    ftWord = 2
End Enum

Property Get GetPathOfReport(eReport As eReport) As String
    Select Case eReport
        Case rArticleReport
            GetPathOfReport = App.Path & "\Reports\ArticlesReport.rpt"
            
        Case rCategoriesReport
            GetPathOfReport = App.Path & "\Reports\CategoriesReport.rpt"
    End Select
End Property

Private Sub Form_Load()
    LoadReportList
    
    'implicit load Report
    cboReports.ListIndex = eReport.rArticleReport
End Sub

Private Sub miSaveAs_Excel_Click()
    Select Case cboReports.ListIndex
        Case eReport.rArticleReport
            ExportReport mReport, App.Path & "\ArticlesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".xls", ftExcel, True
            
        Case eReport.rCategoriesReport
            ExportReport mReport, App.Path & "\CategoriesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".xls", ftExcel, True

        Case Else
            ExportReport mReport, App.Path & "\ArticlesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".xls", ftExcel, True
    End Select
End Sub

Private Sub miSaveAs_PDF_Click()
    Select Case cboReports.ListIndex
        Case eReport.rArticleReport
            ExportReport mReport, App.Path & "\ArticlesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".pdf", ftPDF, True
            
        Case eReport.rCategoriesReport
            ExportReport mReport, App.Path & "\rCategoriesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".pdf", ftPDF, True

        Case Else
            ExportReport mReport, App.Path & "\ArticlesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".pdf", ftPDF, True
    End Select
End Sub

Private Sub miSaveAs_Word_Click()
    Select Case cboReports.ListIndex
        Case eReport.rArticleReport
            ExportReport mReport, App.Path & "\ArticlesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".doc", ftWord, True
            
        Case eReport.rCategoriesReport
            ExportReport mReport, App.Path & "\rCategoriesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".doc", ftWord, True

        Case Else
            ExportReport mReport, App.Path & "\ArticlesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".doc", ftWord, True
    End Select
End Sub

Private Sub cboReports_Click()
    Dim sPathReport As String
    Dim rsData As ADODB.Recordset
    
    Select Case cboReports.ListIndex
        Case eReport.rArticleReport '0
            Set rsData = modArticleHelper.ConvertToRecordset(modSingletonRepository.GetArticleRepository().GetArticles())
            sPathReport = Me.GetPathOfReport(rArticleReport)
            
        Case eReport.rCategoriesReport '1
            Set rsData = modCategoryHelper.ConvertToRecordset(modSingletonRepository.GetCategoryRepository().GetCategories())
            sPathReport = Me.GetPathOfReport(rCategoriesReport)
    End Select
        
    LoadReport mReport, rsData, sPathReport
End Sub

Private Sub LoadReport(ByRef crxReport As CRAXDRT.Report, rsData As ADODB.Recordset, Optional sPathReport As String)
    Static crxApp As New CRAXDRT.Application
    
    If Dir(sPathReport) = "" Then
        MsgBox "File not found: " & sPathReport, vbCritical, "LoadReport - Error"
        Exit Sub
    End If

    Set crxReport = crxApp.OpenReport(sPathReport)
    
    crxReport.Database.SetDataSource rsData
    
    crViewer.ReportSource = crxReport
    crViewer.ViewReport
    crViewer.Zoom ZOOM_FULL_PAGE
End Sub

Private Sub ExportReport(crxReport As CRAXDRT.Report, sFileName As String, eFormatType As eFormatType, Optional bShowDialogbox As Boolean = False)
    Dim crxExportOptions As CRAXDRT.ExportOptions
    
    With crxReport
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = False
    End With
    
    If bShowDialogbox Then
        With dlgSaveAs
            .CancelError = True
            
            Select Case eFormatType
                Case ftPDF
                    .Filter = "PDF Files (*.pdf)|*.pdf"
                    
                Case ftExcel
                    .Filter = "Excel Files (*.xls)|*.xls"
                
                Case ftWord
                    .Filter = "Word Files (*.doc)|*.doc"
            End Select
            
            .DialogTitle = "Save As"
            .FileName = sFileName
            .InitDir = App.Path
            
            On Error Resume Next
            .ShowSave
            
            If Err.Number <> 0 Then
                MsgBox "Export canceled.", vbInformation
                Exit Sub
            End If
            
            sFileName = .FileName
            On Error GoTo 0
        End With
        
        bShowDialogbox = False
    End If
    
    Set crxExportOptions = crxReport.ExportOptions
    
    With crxExportOptions
        .DestinationType = crEDTDiskFile
        .DiskFileName = sFileName
        
        Select Case eFormatType
            Case ftPDF
                .FormatType = crEFTPortableDocFormat
                
            Case ftExcel
                .FormatType = crEFTExcel80
                
            Case ftWord
                .FormatType = crEFTWordForWindows
        End Select
        
        .PDFExportAllPages = True
    End With
    
    crxReport.Export bShowDialogbox
End Sub

Private Sub LoadCombobox(ParamArray vParam() As Variant)
    For i = 0 To UBound(vParam)
        cboReports.AddItem vParam(i)
    Next
End Sub

Private Sub LoadReportList()
    cboReports.AddItem "ArticlesReport.rpt"
    cboReports.AddItem "CategoriesReport.rpt"
End Sub
