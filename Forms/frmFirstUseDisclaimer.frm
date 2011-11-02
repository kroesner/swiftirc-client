VERSION 5.00
Begin VB.Form frmFirstUseDisclaimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SwiftIRC Agreement"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5370
   Icon            =   "frmFirstUseDisclaimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox textAgreement 
      Height          =   1470
      Left            =   225
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmFirstUseDisclaimer.frx":000C
      Top             =   180
      Width           =   2955
   End
End
Attribute VB_Name = "frmFirstUseDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_labelManager As New CLabelManager
Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1
Private WithEvents m_buttonRules As swiftIrc.ctlButton
Attribute m_buttonRules.VB_VarHelpID = -1
Private WithEvents m_textAgree As VB.TextBox
Attribute m_textAgree.VB_VarHelpID = -1

Private m_client As swiftIrc.SwiftIrcClient

Friend Sub init(client As swiftIrc.SwiftIrcClient)
    Set m_client = client
End Sub

Private Sub initControls()
    Set m_buttonOk = addButton(Controls, "&Ok", Me.ScaleWidth - 175, Me.ScaleHeight - 40, 75, 20)
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, 20)
    Set m_buttonRules = addButton(Controls, "&View SwiftIRC rules", 20, Me.ScaleHeight - 40, 120, 20)
    Set m_textAgree = createControl(Controls, "VB.TextBox", "agree")

    m_textAgree.Move (Me.ScaleWidth / 2 - 50), Me.ScaleHeight - 70, 100, 20
    textAgreement.Move 20, 20, Me.ScaleWidth - 40, Me.ScaleHeight - 100
    
    textAgreement.fontName = getBestDefaultFont
    textAgreement.fontSize = 10
    
    textAgreement.text = "By using this IRC client you agree that:" & vbCrLf & vbCrLf _
        & "* You are 13 or more years old" & vbCrLf & vbCrLf _
        & "* You Will not use this software for any illegal or immoral purpose" & vbCrLf & vbCrLf _
        & "* You will remain security conscious when communicating with others over IRC" & vbCrLf & vbCrLf _
        & "* You Will follow the rules outlined at http://www.swiftirc.net/index.php?page=rules" & vbCrLf & vbCrLf _
        & "* You will not hold SwiftIRC responsible for any loss, perceived or otherwise, that you may suffer " _
        & "from your use of this IRC client." & vbCrLf & vbCrLf _
        & "If you agree to these terms, please type ""I agree"" in the field provided below and then press " _
        & "the Ok button"
    
End Sub

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    SetBkMode Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    
    initControls
    updateColours Controls
End Sub

Private Sub Form_Paint()
    SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    
    m_labelManager.renderLabels Me.hdc
End Sub

Private Sub m_buttonCancel_clicked()
    Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
    If StrComp(m_textAgree.text, "I Agree", vbTextCompare) <> 0 Then
        MsgBox "You must enter ""I Agree"" in the field indicated to continue", vbCritical, "Incorrect/missing information"
        Exit Sub
    End If
    
    settings.acceptedFirstUse = True
    settings.saveSettings
    
    Me.Hide
End Sub

Private Sub m_buttonRules_clicked()
    m_client.visitUrl "http://www.swiftirc.net/index.php?page=rules", False
End Sub
