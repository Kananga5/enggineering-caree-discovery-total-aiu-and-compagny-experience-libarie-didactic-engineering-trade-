Attribute VB_Name = "Module1"

Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub ComboBox3_Change()

End Sub

Private Sub ComboBox5_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)

End Sub

Private Sub UserForm_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Layout()

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)

End Sub

Private Sub UserForm_Resize()

End Sub

Private Sub UserForm_Terminate()

End Sub

Private Sub UserForm_Zoom(Percent As Integer)

End Sub
Option Explicit

Private Const SHEET_CASES As String = "Cases"

Private isInitializing As Boolean

'==========================
' Lifecycle
'==========================
End Sub
Private Sub UserForm9_Initialize()
    On Error Resume Next
    isInitializing = True
    
    EnsureCasesSheet
    
    ' Populate top-level lists
    With Me.ComboBox1 ' Case Type
        .Clear
        .AddItem "Refund"
        .AddItem "Compensation"
        .AddItem "Recognition"
        .AddItem "Insurance claim"
    End With
    
    With Me.ComboBox3 ' Issuing Body
        .Clear
        .AddItem "Institution"
        .AddItem "SETA"
        .AddItem "QCTO"
        .AddItem "CCMA"
        .AddItem "Department of Employment and Labour"
        .AddItem "Other"
    End With
    
    With Me.ComboBox5 ' Desired Outcome
        .Clear
        .AddItem "Refund"
        .AddItem "Credit"
        .AddItem "Provisional certificate"
        .AddItem "Appeal"
        .AddItem "Escalation"
        .AddItem "Correction/Letter of completion"
    End With
    
    ' Priority toggle
    Me.OptionButton1.Caption = "Visa/Job critical"
    Me.OptionButton1.Value = False
    
    ' Sensible defaults
    Me.ComboBox1.ListIndex = -1
    Me.ComboBox2.Clear
    Me.ComboBox3.ListIndex = -1
    Me.ComboBox5.ListIndex = -1
    
    isInitializing = False
    On Error GoTo 0
End Sub

Private Sub UserForm9_Activate()
    On Error Resume Next
    If Me.ComboBox1.ListCount > 0 Then Me.ComboBox1.SetFocus
    On Error GoTo 0
End Sub

Private Sub UserForm9_Terminate()
    ' No special teardown
End Sub


'==========================
' Commands
'==========================

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CASES)
    
    Dim r As Long
    r = NextFreeRow(ws)
    
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = caseId
    ws.Cells(r, 3).Value = Nz(Me.ComboBox1.Value)
    ws.Cells(r, 4).Value = Nz(Me.ComboBox2.Value)
    ws.Cells(r, 5).Value = Nz(Me.ComboBox3.Value)
    ws.Cells(r, 6).Value = Nz(Me.ComboBox5.Value)
    ws.Cells(r, 7).Value = IIf(Me.OptionButton1.Value, "High", "Normal")
    ws.Cells(r, 8).Value = "Submitted"
    ws.Cells(r, 9).Value = "" ' Notes (optional)
    
    MsgBox "Case submitted: " & caseId, vbInformation, "Success"
    
    ResetForm
End Sub

()
    ' Save draft (partial allowed)
    Dim caseId As String
    caseId = GenerateCaseId
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CASES)
    
    Dim r As Long
    r = NextFreeRow(ws)
    
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = caseId
    ws.Cells(r, 3).Value = Nz(Me.ComboBox1.Value)
    ws.Cells(r, 4).Value = Nz(Me.ComboBox2.Value)
    ws.Cells(r, 5).Value = Nz(Me.ComboBox3.Value)
    ws.Cells(r, 6).Value = Nz(Me.ComboBox5.Value)
    ws.Cells(r, 7).Value = IIf(Me.OptionButton1.Value, "High", "Normal")
    ws.Cells(r, 8).Value = "Draft"
    ws.Cells(r, 9).Value = "" ' Notes
    
    MsgBox "Draft saved: " & caseId, vbInformation, "Saved"
End Sub

()
    ' Reset
    ResetForm
End Sub

()
    ' Close
    Unload Me
End Sub

'==========================
' User experience events
'==========================
)
    ' ESC closes; Ctrl+S saves draft
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyS And (Shift And fmCtrlMask) = fmCtrlMask Then
        CommandButton2_Click
    End If
End Sub

()
    ' No-op
End Sub

)
    ' No-op
End Sub

()
    ' Hook for responsive layout if needed
End Sub

)
    ' No-op
End Sub

)
End Sub

)
End Sub

)
    ' Keep default behavior
End Sub

()
    ' Optionally reposition/resize controls here
End Sub

'==========================
' Helpers
'==========================
Private Sub FillScenarioList(ByVal caseType As String)
    Me.ComboBox2.Clear
    
    Select Case LCase$(Trim$(caseType))
        Case "refund"
            Me.ComboBox2.AddItem "Training not delivered"
            Me.ComboBox2.AddItem "Material defects / not as described"
            Me.ComboBox2.AddItem "Admin error in registration"
            Me.ComboBox2.AddItem "Overbilling"
        Case "compensation"
            Me.ComboBox2.AddItem "Diploma printing delay (loss of opportunity)"
            Me.ComboBox2.AddItem "Application rejected without due cause"
            Me.ComboBox2.AddItem "Published without registration confirmation"
        Case "recognition"
            Me.ComboBox2.AddItem "Request provisional certificate"
            Me.ComboBox2.AddItem "Request letter of completion"
            Me.ComboBox2.AddItem "Appeal assessment outcome"
        Case "insurance claim"
            Me.ComboBox2.AddItem "Policy claim for learning costs"
            Me.ComboBox2.AddItem "Denied claim appeal"
        Case Else
            ' Generic fallback
            Me.ComboBox2.AddItem "Other"
    End Select
End Sub

Private Sub SuggestOutcome()
    ' Suggest an outcome based on scenario keywords (non-binding)
    Dim s As String
    s = LCase$(Nz(Me.ComboBox2.Value))
    
    If s Like "*not delivered*" Or s Like "*overbilling*" Then
        SelectOutcomeIfExists "Refund"
    ElseIf s Like "*printing*" Or s Like "*provisional*" Or s Like "*completion*" Then
        SelectOutcomeIfExists "Provisional certificate"
    ElseIf s Like "*rejected*" Or s Like "*appeal*" Then
        SelectOutcomeIfExists "Appeal"
    ElseIf s Like "*published*" Or s Like "*admin*" Then
        SelectOutcomeIfExists "Correction/Letter of completion"
    End If
End Sub

Private Sub SelectOutcomeIfExists(ByVal text As String)
    Dim i As Long
    For i = 0 To Me.ComboBox5.ListCount - 1
        If StrComp(Me.ComboBox5.List(i), text, vbTextCompare) = 0 Then
            Me.ComboBox5.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Function ValidateForm(ByVal isFinal As Boolean) As Boolean
    ValidateForm = False
    
    Dim missing As String
    missing = ""
    
    If Len(Trim$(Nz(Me.ComboBox1.Value))) = 0 Then missing = missing & "- Case Type" & vbCrLf
    If Len(Trim$(Nz(Me.ComboBox2.Value))) = 0 Then missing = missing & "- Scenario" & vbCrLf
    If Len(Trim$(Nz(Me.ComboBox3.Value))) = 0 Then missing = missing & "- Issuing Body" & vbCrLf
    
    If isFinal And Len(missing) > 0 Then
        MsgBox "Please complete the following before submitting:" & vbCrLf & vbCrLf & missing, vbExclamation, "Incomplete"
        Exit Function
    End If
    
    ValidateForm = True
End Function

Private Sub ResetForm()
    isInitializing = True
    
    Me.ComboBox1.ListIndex = -1
    Me.ComboBox2.Clear
    Me.ComboBox3.ListIndex = -1
    Me.ComboBox5.ListIndex = -1
    Me.OptionButton1.Value = False
    
    isInitializing = False
End Sub

Private Function GenerateCaseId() As String
    GenerateCaseId = "CASE-" & Format(Now, "yymmdd-hhnnss")
End Function


    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If r < 2 Then
        NextFreeRow = 2
    Else
        NextFreeRow = r + 1
    End If
End Function

Private Sub EnsureCasesSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CASES)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_CASES
    End If
    
    ' Headers if empty
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "DateTime"
        ws.Cells(1, 2).Value = "CaseID"
        ws.Cells(1, 3).Value = "CaseType"
        ws.Cells(1, 4).Value = "Scenario"
        ws.Cells(1, 5).Value = "IssuingBody"
        ws.Cells(1, 6).Value = "DesiredOutcome"
        ws.Cells(1, 7).Value = "Priority"
        ws.Cells(1, 8).Value = "Status"
        ws.Cells(1, 9).Value = "Notes"
    End If
End Sub

Private Function Nz(ByVal v) As String
    If IsNull(v) Then
        Nz = ""
    Else
        Nz = CStr(v)
    End If
End Function

()

End Sub

()

End Sub



End Sub

Private Sub ComboBox4_Change()

End Sub



End Sub



End Sub



End Sub



End Sub



End

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub TextBox1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox1_DropButtonClick()

End Sub

Private Sub TextBox1_Enter()

End Sub

Private Sub TextBox1_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub


End Sub


    Select Case ComboBox1.Value
        Case "Kinematics"
            Label1.Caption = "Select motion type"
        Case "Statics"
            Label1.Caption = "Select force system"
        Case "Dynamics"Private Sub CommandButton1_Click()
    If ComboBox1.Value = "" Or ComboBox2.Value = "" Then
        MsgBox "Please complete all selections", vbExclamation
        Exit Sub
    End If

    ' Example: Generate diagnostic output
    TextBox1.text = "Running simulation for " & ComboBox1.Value & " with parameter " & ComboBox2.Value
End Sub
()
    If Len(TextBox1.text) > 50 Then
        Label2.Caption = "Input exceeds recommended length"
    Else
        Label2.Caption = ""
    End If
End Sub
            Label1.Caption = "Select energy model"
    End Select
End Sub



End Sub


End Sub



End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label13_Click()

End Sub


End Sub

Private Sub Label4_Click()

End Sub



End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub ScrollBar1_Change()

End Sub


End Sub

Private Sub TextBox10_Change()

End Sub

Private Sub TextBox12_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox9_Change()

End Sub

Private Sub ToggleButton1_AfterUpdate()

End Sub

Private Sub ToggleButton1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub ToggleButton1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ToggleButton1_Change()

End Sub

Private Sub ToggleButton1_Click()

End Sub

Private Sub ToggleButton1_Enter()

End Sub

Private Sub ToggleButton1_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub ToggleButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub ToggleButton1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub ToggleButton1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub ToggleButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub ToggleButton1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub


End Sub
()
    MsgBox "Select the diagnostic domain from ListBox1 to proceed.", vbInformation
End Sub
()
    Select Case ListBox1.Value
        Case "Kinematics"
            Frame1.Caption = "Motion Parameters"
        Case "Statics"
            Frame1.Caption = "Force Systems"
        Case "Dynamics"
            Frame1.Caption = "Energy Models"
    End Select
End Sub
()
    If Not IsNumeric(TextBox2.text) Then
        Label3.Caption = "Please enter a numeric value"
    Else
        Label3.Caption = ""
    End If
End Sub
)
    If ToggleButton1.Value = True Then
        Label4.Caption = "Advanced Mode Enabled"
    Else
        Label4.Caption = "Basic Mode Active"
    End If
End Sub
()
    Label5.Caption = "Now viewing: " & MultiPage1.Pages(MultiPage1.Value).Caption
End Sub



End Sub


End Sub


End Sub



End Sub


End Sub


End Sub


End Sub



End Sub

Private Sub TextBox11_Change()

End Sub



End Sub

Private Sub TextBox13_Change()

End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox15_Change()

End Sub

Private Sub TextBox16_Change()

End Sub

Private Sub TextBox17_Change()

End Sub

Private Sub TextBox19_Change()

End Sub

Private Sub TextBox20_Change()

End Sub

Private Sub TextBox21_Change()

End Sub

Private Sub TextBox23_Change()

End Sub



End Sub

Private Sub TextBox5_Change()

End Sub


End Sub


End Sub



End Sub


End Sub



End Sub

Private Sub UserForm_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub



End Sub



End Sub



End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub



End Sub


End Sub


End Sub



End Sub

Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)

End Sub


End Sub

    If Not IsNumeric(TextBox12.text) Then
        Label6.Caption = "Voltage must be numeric"
    Else
        Label6.Caption = ""
    End If
End Sub
()
    If TextBox12.text = "" Or TextBox13.text = "" Then
        MsgBox "Please enter all required parameters", vbExclamation
        Exit Sub
    End If

    Dim Voltage As Double, resistance As Double
    Voltage = CDbl(TextBox12.text)
    resistance = CDbl(TextBox13.text)

    TextBox14.text = "Current: " & Format(Voltage / resistance, "0.00") & " A"
End Sub
)
    MsgBox "Enter voltage in volts and resistance in ohms to compute current.", vbInformation
End Sub



End Sub


End Sub



End Sub

Private Sub SpinButton1_Change()

End Sub

Private Sub SpinButton2_Change()

End Sub

Private Sub TabStrip1_Change()

End Sub


End Sub



End Sub



End Sub


End Sub

