VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListViewUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents mListView As ListView
Attribute mListView.VB_VarHelpID = -1
Private mColumnSortOrder
Private mSortColumn As Integer
Private mSortedAsc As Boolean

Public Sub Initialize(ByRef lvw As ListView)
    Set mListView = lvw
End Sub

Private Sub mListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub mListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mSortedAsc Then
        mSortedAsc = False
    Else
        mSortedAsc = True
    End If
    
    Select Case ColumnHeader.Text
        Case "Id", "Articles"
            If mSortedAsc = True Then
                Call SortColumn(mListView, ColumnHeader.Index, sortDescending, sortNumeric)
            Else
                Call SortColumn(mListView, ColumnHeader.Index, sortAscending, sortNumeric)
            End If
            
        Case "Name", "Details", "Category"
            If mSortedAsc = True Then
                Call SortColumn(mListView, ColumnHeader.Index, sortDescending, sortAlpha)
            Else
                Call SortColumn(mListView, ColumnHeader.Index, sortAscending, sortAlpha)
            End If
    End Select
End Sub
