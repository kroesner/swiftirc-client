VERSION 5.00
Begin VB.Form frmGenericInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Title"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3555
   Icon            =   "frmGenericInput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   237
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGenericInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1
Private m_textInput As VB.TextBox

Private m_labelManager As New CLabelManager

Private m_title As String
Private m_caption As String
Private m_default As String
Private m_cancelled As Boolean

Public Property Get cancelled() As Boolean
    cancelled = m_cancelled
End Property

Public Property Get value() As String
    value = m_textInput.text
End Property

Public Sub init(title As String, caption As String, Optional default As String = vbNullString, Optional password As Boolean = False)
    m_title = title
    m_caption = caption
    
    Me.caption = title
    m_labelManager.addLabel caption, ltNormal, 30, 20
    m_default = default
    m_textInput.text = m_default
    
    If password Then
        m_textInput.PasswordChar = "*"
    End If
End Sub

Private Sub Form_Activate()
    m_textInput.setFocus
End Sub

Private Sub Form_Initialize()
    initControls
    updateColours Controls
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        m_buttonCancel_clicked
    End If
End Sub

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    updateColours Controls
End Sub

Private Sub initControls()
    Set m_buttonOk = addButton(Controls, "&OK", Me.ScaleWidth - 165, Me.ScaleHeight - 40, 70, 20)
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 90, Me.ScaleHeight - 40, 70, 20)
    Set m_textInput = createControl(Controls, "VB.TextBox", "textInput")
    
    m_textInput.Move 30, 40, Me.ScaleWidth - 60, 20
End Sub

Private Sub Form_Paint()
    SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    
    m_labelManager.renderLabels Me.hdc
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    m_cancelled = True
End Sub

Private Sub m_buttonCancel_clicked()
    m_cancelled = True
    Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
    Me.Hide
End Sub
