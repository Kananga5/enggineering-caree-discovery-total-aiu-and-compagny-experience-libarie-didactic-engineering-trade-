VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm13 
   Caption         =   "UserForm13"
   ClientHeight    =   9816
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   19488
   OleObjectBlob   =   "UserForm9991prog licens.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Frame1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub Frame1_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub Frame1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub Frame1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub Frame1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub Frame1_Layout()

End Sub

Private Sub Frame1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub Frame1_RemoveControl(ByVal Control As MSForms.Control)

End Sub

Private Sub Frame1_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub
Public Function GenerateSHA256(ByVal inputText As String) As String
    Dim shaObj As CSHA256
    Set shaObj = New CSHA256
    GenerateSHA256 = shaObj.SHA256(inputText)
    Set shaObj = Nothing
End Function

    Dim productName As String
    productName = TextBox1.text
    TextBox2.text = GenerateSHA256(productName) ' SHA ID output
End Sub
 ' "Issue Certificate" button
    If TextBox2.text <> "" Then
        MsgBox "Certificate issued for product: " & TextBox1.text & vbCrLf & "SHA ID: " & TextBox2.text
        ' Optional: Log to registry or export to file
    Else
        MsgBox "SHA ID missing. Cannot issue certificate."
    End If
End Sub


