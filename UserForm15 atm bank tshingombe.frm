VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm15 
   Caption         =   "UserForm15"
   ClientHeight    =   10536
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   20184
   OleObjectBlob   =   "UserForm15 atm bank tshingombe.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub CommandButton5_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox7_Change()

End Sub

Private Sub TextBox8_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox8_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox8_DropButtonClick()

End Sub

Private Sub TextBox8_Enter()

End Sub

Private Sub TextBox8_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub TextBox8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub TextBox8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_Click()

    Public Class transactionsGBox
    
    
    
    
    
    
    
    
    

        Const SERVICE_CHARGE_DECIMAL As Decimal = 6.5
    
    
    
    
    
    

        Const PIN As Integer = 9343
    
    
    
    
    
    
    
    
    

        Dim Balance As Decimal = 150
    
    
    
    
    
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    

        Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
    

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles topUpButton.CheckedChanged

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub transactionsGBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Function withdraw(ByVal amount As Decimal)
    
    
    
    
    
    

            Balance -= amount
    
    
    
    
    
    
    
    
    

            Return Balance
    
    
    
    
    
    
    
    
    
    

        End Function
    
    
    
    
    
    
    
    
    
    

        Private Function deposit(ByRef amount As Decimal)
    
    
    
    
    
    

            Balance += amount
    
    
    
    
    
    
    
    
    

            Return Balance
    
    
    
    
    
    
    
    
    
    

        End Function
    
    
    
    
    
    
    
    
    
    

        Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clearButton1.Click

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles confirmButton.Click

            If pinBox.Text = "9343" Then
    
    
    
    
    
    
    
    

                transactionGroupBox.Enabled = True
    
    
    
    
    
    
    

                previewButton.Enabled = True
    
    
    
    
    
    
    
    

                proceedButton.Enabled = True
    
    
    
    
    
    
    
    

                pinBox.Enabled = False
    
    
    
    
    
    
    
    
    

            Else
    
    
    
    
    
    
    
    
    
    
    

                MessageBox.Show("Incorrect pin, try again", "Pin Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If
    
    
    
    
    
    
    
    
    
    

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
    

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clearButton2.Click

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub exitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exitButton.Click

            Me.Close()
    
    
    
    
    
    
    
    
    
    

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub previewButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles previewButton.Click

            If depositButton.Checked = True Then
    
    
    
    
    
    
    

                previewBalance.Text = deposit(transactionValueBox.Text)
    
    
    
    
    

            Else
    
    
    
    
    
    
    
    
    
    
    

                previewBalance.Text = withdraw(transactionValueBox.Text)
    
    
    
    
    

            End If
    
    
    
    
    
    
    
    
    
    

        End Sub
    
    
    
    
    
    
    
    
    
    

        Private Sub proceedButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles proceedButton.Click

            If depositButton.Checked = True Then
    
    
    
    
    
    
    

                finalBalance.Text = deposit(transactionValueBox.Text)
    
    
    
    
    
    

            Else
    
    
    
    
    
    
    
    
    
    
    

                finalBalance.Text = withdraw(transactionValueBox.Text)
    
    
    
    
    
    

            End If
    
    
    
    
    
    
    
    
    
    

        End Sub
    
    
    
    
    
    
    
    
    
    

    End Class
    
    
    

End Sub

