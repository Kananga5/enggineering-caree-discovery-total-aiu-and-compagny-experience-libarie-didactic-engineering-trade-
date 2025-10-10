VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

End Sub

Private Sub Workbook_PivotTableCloseConnection(ByVal Target As PivotTable)

End Sub

Private Sub Workbook_PivotTableOpenConnection(ByVal Target As PivotTable)

End Sub

Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)

End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

End Sub

Private Sub Workbook_SheetCalculate(ByVal Sh As Object)

End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)

End Sub

Private Sub Workbook_SheetLensGalleryRenderComplete(ByVal Sh As Object)

End Sub

Private Sub Workbook_SheetPivotTableBeforeAllocateChanges(ByVal Sh As Object, ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean)

End Sub

Private Sub Workbook_SheetPivotTableBeforeDiscardChanges(ByVal Sh As Object, ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long)

End Sub

Private Sub Workbook_SheetPivotTableChangeSync(ByVal Sh As Object, ByVal Target As PivotTable)

End Sub

Private Sub Workbook_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)

End Sub

Private Sub Workbook_SheetTableUpdate(ByVal Sh As Object, ByVal Target As TableObject)

End Sub

Private Sub Workbook_Sync(ByVal SyncEventType As Office.MsoSyncEventType)

End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)

End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)

End Sub




    Call InitializeCredentialEngine
    Call LoadRubricTemplates
    MsgBox "Credential system initialized.", vbInformation
End Sub



    If Sh.Name = "RubricScores" Then
        Call RecalculateDiagnostics(Sh)
    End If
End Sub

    LogAction "Attempted to delete sheet: " & Sh.Name
    If Sh.Name = "CredentialMap" Then
        MsgBox "Deletion blocked: Credential map is protected.", vbCritical
        Cancel = True
    End If
End Sub

    Select Case SyncEventType
        Case msoSyncEventDownloadInitiated
            MsgBox "Syncing latest credential data from cloud..."
        Case msoSyncEventUploadInitiated
            MsgBox "Uploading updated learner portfolios..."
    End Select
End Sub

