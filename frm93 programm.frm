VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm13 
   Caption         =   "UserForm13"
   ClientHeight    =   9804
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   20100
   OleObjectBlob   =   "frm93 programm.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frm13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Frame1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

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

Private Sub TextBox1_Change()

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

Private Sub UserForm_Click()

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



