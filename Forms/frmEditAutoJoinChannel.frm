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
    success = m_success
End Property

Public Property Get channel() As String
    channel = m_fieldChannel.value
End Property

Public Property Let channel(newValue As String)
    m_fieldChannel.value = newValue
End Property

Public Property Get key() As String
    key = m_fieldKey.value
End Property

Public Property Let key(newValue As String)
    m_fieldKey.value = newValue
End Property

Private Sub Form_Initialize()
    Set m_fieldChannel = addField(Controls, "Channel name:", 20, 20, Me.ScaleWidth - 40, 20)
    Set m_fieldKey = addField(Controls, "Key/password:", 20, 45, Me.ScaleWidth - 40, 20)
    Set m_buttonSave = addButton(Controls, "Save", 20, 75, 100, 20)
    Set m_buttonCancel = addButton(Controls, "Cancel", 130, 75, 100, 20)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Me.Hide
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        m_buttonSave_clicked
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
End Sub

Private Sub Form_Paint()
    SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
End Sub

Private Sub m_buttonCancel_clicked()
    Me.Hide
End Sub

Private Sub m_buttonSave_clicked()
    If LenB(m_fieldChannel.value) = 0 Then
        MsgBox "Please enter a channel name", vbCritical, "Missing fields"
        Exit Sub
    End If

    m_success = True
    Me.Hide
End Sub
