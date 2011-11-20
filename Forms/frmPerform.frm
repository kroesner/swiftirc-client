VERSION 5.00
Begin VB.Form frmPerform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit perform"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmPerform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox textPerform 
      Height          =   1635
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2475
   End
End
Attribute VB_Name = "frmPerform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_labelManager As New CLabelManager
Private m_checkEnablePerform As VB.CheckBox
Private WithEvents m_buttonSave As ctlButton
Attribute m_buttonSave.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1

Private m_enablePerform As Boolean
Private m_perform As String

Private m_cancel As Boolean

Public Property Get cancelled() As Boolean
10        cancelled = m_cancel
End Property

Public Property Get perform() As String
10        perform = textPerform.text
End Property

Public Property Let perform(newValue As String)
10        m_perform = newValue
End Property

Public Property Get enablePerform() As Boolean
10        enablePerform = -m_checkEnablePerform.value
End Property

Public Property Let enablePerform(newValue As Boolean)
10        m_enablePerform = newValue
End Property

Private Sub Form_Load()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
20        initControls
30        updateColours Controls
End Sub

Private Sub initControls()
10        m_labelManager.addLabel "Reference your own nickname with $me", ltNormal, 20, Me.ScaleHeight - 50

20        textPerform.Move 20, 40, Me.ScaleWidth - 40, Me.ScaleHeight - 95
          
30        Set m_checkEnablePerform = addCheckBox(Controls, "Enable perform", 20, 15, 125, 20)
          
40        Set m_buttonSave = addButton(Controls, "&Save", Me.ScaleWidth - 175, Me.ScaleHeight - 35, 75, 20)
50        Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 35, 75, 20)
          
60        m_checkEnablePerform.value = -m_enablePerform
70        textPerform.text = m_perform
End Sub

Private Sub Form_Paint()
10        SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
          
30        SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
40        m_labelManager.renderLabels Me.hdc
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
10        m_cancel = True
End Sub

Private Sub m_buttonCancel_clicked()
10        m_cancel = True
20        Me.Hide
End Sub

Private Sub m_buttonSave_clicked()
10        Me.Hide
End Sub


