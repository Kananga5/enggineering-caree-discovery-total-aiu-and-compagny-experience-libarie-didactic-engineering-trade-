VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "UserForm5"
   ClientHeight    =   9984
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   20136
   OleObjectBlob   =   "UserForm90000,,inventory.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub TabStrip1_Change()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox10_Change()

End Sub

Private Sub TextBox11_Change()

End Sub

Private Sub TextBox12_Change()

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

Private Sub TextBox18_Change()

End Sub

Private Sub TextBox19_Change()

End Sub

Private Sub TextBox20_Change()

End Sub

Private Sub TextBox21_Change()

End Sub

Private Sub TextBox22_Change()

End Sub

Private Sub TextBox23_Change()

End Sub

Private Sub TextBox24_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox9_Change()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)

End Sub

Private Sub UserForm_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_Resize()

End Sub

Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)

End Sub

Private Sub UserForm_Terminate()

End Sub

Private Sub UserForm_Zoom(Percent As Integer)

    Dim s As Worksheet: Set s = WS("Simulations")
    Dim lastRow As Long: lastRow = s.Cells(s.rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 2 To lastRow
        If Not IsEmpty(s.Cells(i, 2).Value) Then
            cmbLearnerID.AddItem s.Cells(i, 2).Value
        End If
    Next i
End Sub

Private Sub btnValidate_Click()
    Dim learnerID As String: learnerID = cmbLearnerID.Value
    If learnerID = "" Then MsgBox "Please select a learner ID.": Exit Sub

    If PortfolioGatesOK(learnerID) Then
        lblSimStatus.Caption = "Portfolio Valid ?"
        lblSimStatus.ForeColor = vbGreen
    Else
        lblSimStatus.Caption = "Portfolio Invalid ?"
        lblSimStatus.ForeColor = vbRed
    End If
End Sub
Private Sub btnDecayCalc_Click()
    Dim C0 As Double, lambda As Double, t As Double
    C0 = CDbl(txtC0.Value)
    lambda = CDbl(txtLambda.Value)
    t = CDbl(txtTime.Value)
    MsgBox "C(t) = " & Format(Decay_C(C0, lambda, t), "0.000")
End Sub


