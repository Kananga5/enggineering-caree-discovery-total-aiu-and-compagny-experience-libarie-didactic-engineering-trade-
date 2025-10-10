VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm14 
   Caption         =   "UserForm14"
   ClientHeight    =   10068
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   20112
   OleObjectBlob   =   "UserForm95.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Frame1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub SpinButton1_Change()

End Sub

Private Sub SpinButton2_Change()

End Sub

Private Sub TabStrip1_Change()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox11_Change()

End Sub

Private Sub TextBox13_Change()

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

Private Sub TextBox22_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox9_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox9_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox9_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox9_Change()

End Sub

Private Sub TextBox9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox9_DropButtonClick()

End Sub

Private Sub TextBox9_Enter()

End Sub

Private Sub TextBox9_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub TextBox9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub TextBox9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub TextBox9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_Click()

End Sub

()
    Select Case MultiPage1.Value
        Case 0
            Label1.Caption = "Electrical Simulation Module"
        Case 1
            Label1.Caption = "Credential Mapping Module"
        Case 2
            Label1.Caption = "Portfolio Artifact Builder"
    End Select
End Sub

Private Sub Label10_Click()
    MsgBox "Ensure voltage and resistance inputs match rubric thresholds.", vbInformation
End Sub
Private Sub ListBox3_Click()
    TextBox1.text = ListBox3.Value
    Label6.Caption = "Rubric Level: " & ListBox3.Value
End Sub
()
    If Not IsNumeric(TextBox1.text) Then
        Label7.Caption = "?? Input must be numeric"
    Else
        Label7.Caption = ""
    End If
End Sub
()
    ComboBox1.Value = ""
    ComboBox2.Value = ""
    ListBox1.Clear
    ListBox3.Clear
    TextBox1.text = ""
    Label1.Caption = "Diagnostic Interface Ready"
End Sub
