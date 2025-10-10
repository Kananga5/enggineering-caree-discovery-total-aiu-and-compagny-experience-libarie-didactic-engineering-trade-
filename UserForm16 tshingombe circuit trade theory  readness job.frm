VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm16 
   Caption         =   "UserForm16"
   ClientHeight    =   9840
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   20184
   OleObjectBlob   =   "UserForm16 tshingombe circuit trade theory  readness job.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub MultiPage2_Change()

End Sub

Private Sub ScrollBar1_Change()

End Sub

Private Sub SpinButton1_Change()

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

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox7_Change()

End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox9_Change()

End Sub

Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)

End Sub

Private Sub UserForm_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub UserForm_Layout()

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

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
Function K_Rdiv1(R1, R2)
   ' Gain of resistor divider
   K_Rdiv1 = R2 / (R2 + R1)

End FunctionFunction Tri_Wave(t, V1, V2, T1, T2)

' *************************************************************
' Generate Triangle Wave
'
' t - time
' V1 - voltage level 1 (initial voltage)
' V2 - voltage level 2
' T1 - period ramping from V1 to V2
' T2 - period ramping from V2 to V1
'***************************************************************

Dim t_tri, dV_dt1, dV_dt2 As Double
Dim N As Single

' Calculate voltage rates of change (slopes) during T1 and T2
dV_dt1 = (V2 - V1) / T1
dV_dt2 = (V1 - V2) / T2

' given t, how many full cycles have occurred
N = Application.WorksheetFunction.Floor(t / (T1 + T2), 1)

' calc the time point in the current triangle wave
t_tri = t - (T1 + T2) * N

' if during T1, calculate triangle value using V1 and dV_dt1
If t_tri <= T1 Then
    Tri_Wave = V1 + dV_dt1 * t_tri

' if during T2, calculate triangle value using V2 and dV_dt2
Else
   Tri_Wave = V2 + dV_dt2 * (t_tri - T1)

End If
 given t, how many full cycles have occured
N = Application.WorksheetFunction.Floor(t / (T1 + T2), 1)

' calc the time point in the current triangle wave
t_tri = t - (T1 + T2) * N

End FunctionIf t_tri <= T1 ThenElse
   Tri_Wave = V2 + dV_dt2 * (t_tri - T1)
    Tri_Wave = V1 + dV_dt1 * t_tri
    Function K_op_non(R1, R2)
   ' Op amp closed loop gain - non-inverting amplifier
   K_op_non = (R2 + R1) / R1

End Function

Function SineWave(t, Vp, fo, Phase, Vdc)
  ' create sine wave
  ' phase in deg

  Dim pi As Double
  pi = 3.1415927

  'Calc sine wave
  SineWave = Vp * Sin(2 * pi * fo * t + Phase * pi / 180) + Vdc

End Function
 
Function K_op_inv(R1, R2)
   ' Op amp closed loop gain - inverting amplifier
   K_op_inv = -R2 / R1

End Functionn

    



