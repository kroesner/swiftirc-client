VERSION 5.00
Begin VB.Form frmEditAutoJoinChannel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit channel"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3765
   Icon            =   "frmEditAutoJoinChannel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditAutoJoinChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_fieldChannel As ctlField
Private m_fieldKey As ctlField
Private WithEvents m_buttonSave As ctlButton
Attribute m_buttonSave.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1

Private m_success As Boolean

Public Property Get success() As Boolean
10        success = m_success
End Property

Public Property Get channel() As String
10        channel = m_fieldChannel.value
End Property

Public Property Let channel(newValue As String)
10        m_fieldChannel.value = newValue
End Property

Public Property Get key() As String
10        key = m_fieldKey.value
End Property

Public Property Let key(newValue As String)
10        m_fieldKey.value = newValue
End Property

Private Sub Form_Initialize()
10        Set m_fieldChannel = addField(Controls, "Channel name:", 20, 20, Me.ScaleWidth - 40, 20)
20        Set m_fieldKey = addField(Controls, "Key/password:", 20, 45, Me.ScaleWidth - 40, 20)
30        Set m_buttonSave = addButton(Controls, "Save", 20, 75, 100, 20)
40        Set m_buttonCancel = addButton(Controls, "Cancel", 130, 75, 100, 20)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
10        If KeyAscii = vbKeyEscape Then
20            Me.Hide
30            KeyAscii = 0
40        ElseIf KeyAscii = vbKeyReturn Then
50            m_buttonSave_clicked
60            KeyAscii = 0
70        End If
End Sub

Private Sub Form_Load()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
End Sub

Private Sub Form_Paint()
10        SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
End Sub

Private Sub m_buttonCancel_clicked()
10        Me.Hide
End Sub

Private Sub m_buttonSave_clicked()
10        If LenB(m_fieldChannel.value) = 0 Then
20            MsgBox "Please enter a channel name", vbCritical, "Missing fields"
30            Exit Sub
40        End If

50        m_success = True
60        Me.Hide
End Sub
