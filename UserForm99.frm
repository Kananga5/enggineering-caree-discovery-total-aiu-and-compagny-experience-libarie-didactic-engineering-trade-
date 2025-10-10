VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6948
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17760
   OleObjectBlob   =   "UserForm99.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub ComboBox3_Change()

End Sub

Private Sub ComboBox4_Change()

End Sub

Private Sub ComboBox5_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Frame4_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub
()
    Select Case ComboBox1.Value
        Case "Electrical Simulation"
            Label1.Caption = "Domain: Electrical"
        Case "Portfolio Builder"
            Label1.Caption = "Domain: Portfolio"
        Case "Rubric Mapping"
            Label1.Caption = "Domain: Rubric"
    End Select
End Sub
()
    Select Case ComboBox2.Value
        Case "Level 1"
            Label2.Caption = "Rubric: Foundational"
        Case "Level 2"
            Label2.Caption = "Rubric: Intermediate"
        Case "Level 3"
            Label2.Caption = "Rubric: Advanced"
    End Select
End Sub
()
    If Not IsNumeric(TextBox1.text) Then
        Label3.Caption = "?? Enter a numeric value"
    Else
        Label3.Caption = ""
    End If
End Sub
()
    Dim inputVal As Double
    inputVal = Val(TextBox1.text)

    If inputVal = 0 Then
        MsgBox "Input cannot be zero", vbCritical
        Exit Sub
    End If

    Label1.Caption = "Diagnostic Result: " & Format(inputVal * 1.5, "0.00")
End Sub
()
    MsgBox "Artifact exported to credential folder", vbInformation
End Sub
Private Sub CommandButton3_Click()
    MsgBox "Submission successful. Awaiting rubric validation.", vbInformation
End Sub
Private Sub CommandButton4_Click()
    ComboBox1.Value = ""
    ComboBox2.Value = ""
    TextBox1.text = ""
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
End Sub
