VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports"
   ClientHeight    =   11412
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   9852
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11412
   ScaleWidth      =   9852
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExportPDF 
      Caption         =   "Export to PDF"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   10920
      Width           =   1452
   End
   Begin VB.ComboBox cboReports 
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
      Left            =   5040
      TabIndex        =   2
      Text            =   "---Seleccionar---"
      Top             =   10920
      Width           =   4692
   End
   Begin CRVIEWERLibCtl.CRViewer crViewer 
      Height          =   10692
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9612
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
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3480
      TabIndex        =   1
      Top             =   10920
      Width           =   1452
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim m_crxReport As CRAXDRT.Report

Const ZOOM_FULL_WIDTH As Integer = 1
Const ZOOM_FULL_PAGE As Integer = 2

Enum ReportType
    ArticleReport = 1
End Enum

Property Get GetPathOfReport(eReportType As ReportType) As String
    Select Case eReportType
        Case ArticleReport
            GetPathOfReport = App.Path & "\Reports\ArticlesReport.rpt"
    End Select
End Property

Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    Set rsData = modArticleHelper.ConvertToRecordset(modSingletonRepository.GetArticleRepository().GetArticles())
    
    LoadReport m_crxReport, rsData
End Sub

Private Sub cmdExportPDF_Click()
    ExportToPDF m_crxReport, App.Path & "\ArticlesReport_" & Format(Now, "ddmmyyyy_hhmmss") & ".pdf"
End Sub

Private Sub LoadReport(ByRef crxReport As CRAXDRT.Report, rsData As ADODB.Recordset)
    Dim crxApp As New CRAXDRT.Application
    Dim sPathReport As String
    
    sPathReport = Me.GetPathOfReport(ArticleReport)
    
    If Dir(sPathReport) = "" Then
        MsgBox "File not found: " & sPathReport, vbCritical, "LoadReport - Error"
        Exit Sub
    End If
    
    Set crxReport = crxApp.OpenReport(Me.GetPathOfReport(ArticleReport))
    
    crxReport.Database.SetDataSource rsData
    
    crViewer.ReportSource = crxReport
    crViewer.ViewReport
    crViewer.Zoom ZOOM_FULL_PAGE
End Sub

Private Sub ExportToPDF(crxReport As CRAXDRT.Report, sFileName As String)
    If crxReport Is Nothing Then Exit Sub

    Dim crxExportOptions As CRAXDRT.ExportOptions
    
    With crxReport
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = False '*
    End With
    
    Set crxExportOptions = crxReport.ExportOptions
    
    With crxExportOptions
        .DestinationType = crEDTDiskFile
        .DiskFileName = sFileName
        .FormatType = crEFTPortableDocFormat
        .PDFExportAllPages = True
    End With
    
    crxReport.Export False
End Sub
