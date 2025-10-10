VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   9900
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   20148
   OleObjectBlob   =   "UserForm9994 trade theory.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label11_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Label20_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label26_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label31_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub TextBox29_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox31_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox9_Change()

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
If OK = True Then FORM

End Sub
If CANCELL = FALS Then FORM
End Sub
Else: Show
Next FORM
End Sub
If text = True Then



Private Sub Frame1_Click()

End Sub



End Sub



End Sub



End Sub


End Sub






End Sub
If OK = True Then FORM

End Sub
If CANCELL = FALS Then FORM
End Sub
Else: Show
Next FORM
End Sub
If text = True Then

()
    If TextBox9.text = "IEC61850" Then
        MsgBox "Protocol accepted. Proceed to IED configuration."
    ElseIf TextBox9.text = "FDR-TRP" Then
        MsgBox "Feeder tripped. Initiate fault isolation."
    End If
End Sub
If OK = True Then
    MsgBox "Form submitted. Proceed to next phase."
ElseIf Cancel = False Then
    MsgBox "Form cancelled. Restart required."
End If





    ' Capture user input for fault code
    If TextBox9.text = "FDR-TRP" Then
        MsgBox "Feeder tripped. Check relay settings and breaker status."
    End If
End Sub


    ' Log keypress for rubric tracking
    Debug.Print "Key pressed: " & KeyCode
End Sub






P
End Sub


End Sub



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

    






End Sub

Private Sub UserForm17_Terminate()

End Sub

End Subtshingombe fiston
    
Jul 23, 2025, 3:10 PM (2 days ago)
    
to me
Qeios
Peer-approved Preprints Archive

    About
    Ethics
    Plans
    Sign Up Free
    Log in



Views

4,047
Downloads

314
Peer Reviewers

29
Citations

0
Article has an altmetric score of 2
Make Action
PDF
Field

Computer Science
Subfield

Information Systems
Open Peer Review
Preprint

2.79 | 29 peer reviewers
Research Article Dec 11, 2023
https://doi.org/10.32388/JGU5FH
Web-Based Crime Management System for Samara City Main Police Station

Demelash Lemmi Ettisa1, Minota Milkias2
Abstract

Crime is a human experience, and it must be controlled. The Samara town police station plays a significant role in controlling crime. However, the management of crime activities is done manually, which is due to the lack of an automated system that supports the station workers in communicating with citizens to share information and store, retrieve, and manage crime activities. To control crime efficiently, we need to develop online crime management systems.

This project, entitled "Web-Based Crime Management System," is designed to develop an online application in which any citizen can report crimes; if anybody wants to file a complaint against crimes, they must enjoy online communication with the police. This project provides records of crimes that have led to disciplinary cases in addition to being used to simply retrieve information from the database. The system implemented is a typical web-based crime record management system based on client-server architecture, allowing data storage and crime record interchange with police stations.

Corresponding author: Demelash Lemmi Ettisa, nicemanyes@su.edu.et
Chapter One
1. Introduction to the Study

The "Crime Management System" is a web-based website for online complaining and computerized management of crime records (Khan et al., 2008).

A criminal is a popular term used for a person who has committed a crime or has been legally convicted of a crime. "Criminal" also means being connected with a crime. When certain acts or people are involved in or related to a crime, they are termed as criminal (Wex, 2023).

Samara City 's main police station is located in Samara City, within the Afar Regional State. It was established in 1984 E.C. with the purpose of protecting local communities from criminal activities. The Samara City police station is situated near the diesel suppliers in Samara City. In the first phase, there was a small number of police members, including commanders, inspectors, and constables. But recently, more than 170 police members have been employed. It is a well-organized police station that serves in crime prevention; the detection and conviction of criminals depend on a highly responsive manner. The effectiveness of this station is based on how efficient, reliable, and fast it is. As a consequence, the station maintains a large volume of information. To manage their information requirements, the station is currently using an information system. This system is manual and paper-based, where information is passed hand-to-hand, and information is kept in hard-copy paper files stored ordinarily in fili













