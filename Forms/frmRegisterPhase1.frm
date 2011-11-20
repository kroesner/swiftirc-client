VERSION 5.00
Begin VB.Form frmRegisterPhase1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nickname registration"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "frmRegisterPhase1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegisterPhase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_labelManager As New CLabelManager
Private m_fieldNickname As swiftIrc.ctlField
Private m_fieldPassword As swiftIrc.ctlField
Private m_fieldPasswordRepeat As swiftIrc.ctlField
Private m_fieldEmail As swiftIrc.ctlField
Private m_fieldEmailRepeat As swiftIrc.ctlField

Private m_checkTosAgree As VB.CheckBox
Private m_checkSecurePassword As VB.CheckBox

Private WithEvents m_registerSession As CSession
Attribute m_registerSession.VB_VarHelpID = -1

Private m_labelStatus As CLabel

Private WithEvents m_buttonRegister As swiftIrc.ctlButton
Attribute m_buttonRegister.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1
Private WithEvents m_buttonRules As swiftIrc.ctlButton
Attribute m_buttonRules.VB_VarHelpID = -1

Private m_client As swiftIrc.SwiftIrcClient

Friend Sub init(client As swiftIrc.SwiftIrcClient)
10        Set m_client = client
End Sub

Private Sub Form_Initialize()
10        initControls
20        updateColours Controls
End Sub

Private Sub Form_Load()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
20        SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
End Sub

Private Sub initControls()
10        m_labelManager.addLabel "Register a nickname on SwiftIRC", ltHeading, 20, 15
20        Set m_labelStatus = m_labelManager.addLabel("To register a nickname, enter your desired nickname and password along with a valid e-mail address.", ltNormal, 0, 0)
          
30        m_labelStatus.setRect makeRect(20, Me.ScaleWidth - 20, 35, 75)
          
40        Set m_fieldNickname = addField(Controls, "Desired nickname", 20, 90, 300, 20)
50        Set m_fieldPassword = addField(Controls, "Password", 20, 125, 300, 20)
60        Set m_fieldPasswordRepeat = addField(Controls, "Repeat password", 20, 150, 300, 20)
70        Set m_fieldEmail = addField(Controls, "E-mail", 20, 185, 300, 20)
80        Set m_fieldEmailRepeat = addField(Controls, "Repeat e-mail", 20, 210, 300, 20)
          
90        Set m_checkTosAgree = addCheckBox(Controls, "I agree to follow the SwiftIRC rules concerning nickname registration", 20, 240, 300, 25)
100       Set m_checkSecurePassword = addCheckBox(Controls, "I have NOT used my RuneScape password or any of my other important passwords", 20, 270, 300, 25)
          
110       Set m_buttonRegister = addButton(Controls, "&Register", Me.ScaleWidth - 175, Me.ScaleHeight - 40, 75, 20)
120       Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, 20)
          
130       Set m_buttonRules = addButton(Controls, "&View SwiftIRC Rules", 20, Me.ScaleHeight - 40, 125, 20)
          
140       m_fieldNickname.mask = fmIrcNickname
          
150       m_fieldNickname.required = True
160       m_fieldPassword.required = True
170       m_fieldPasswordRepeat.required = True
180       m_fieldEmail.required = True
190       m_fieldEmailRepeat.required = True
          
200       m_fieldPassword.password = True
210       m_fieldPasswordRepeat.password = True
End Sub

Private Sub Form_Paint()
10        SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
          
30        m_labelManager.renderLabels Me.hdc
End Sub

Private Sub m_buttonCancel_clicked()
10        Me.Hide
End Sub

Private Sub m_buttonRegister_clicked()
10        If LenB(m_fieldNickname.value) = 0 Then
20            MsgBox "Please enter a nickname", vbCritical, "Missing information"
30            Exit Sub
40        End If
          
50        If LenB(m_fieldPassword.value) = 0 Then
60            MsgBox "Please enter a password for your nickname", vbCritical, "Missing information"
70            Exit Sub
80        End If
          
90        If m_fieldPassword.value <> m_fieldPasswordRepeat.value Then
100           MsgBox "The passwords you entered did not match", vbCritical, "Mismatch"
110           Exit Sub
120       End If
          
130       If Not m_fieldEmail.value Like "*@*.*" Then
140           MsgBox "Please enter a valid e-mail address", vbCritical, "Invalid input"
150           Exit Sub
160       End If
          
170       If m_fieldEmail.value <> m_fieldEmailRepeat.value Then
180           MsgBox "The e-mail addresses you entered did not match", vbCritical, "Mismatch"
190           Exit Sub
200       End If
          
          Dim session As CSession
          
210       Set session = m_client.findSwiftIRCSession
          
220       If Not session Then
230           Set session = m_client.newSession
240           session.primaryNickname = m_fieldNickname.value
250           session.serverHost = "irc.swiftirc.net"
260           session.serverPort = 6667
270           session.connect
280       End If
End Sub

Private Sub m_buttonRules_clicked()
10        m_client.visitUrl "http://www.swiftirc.net/index.php?page=rules", False
End Sub
