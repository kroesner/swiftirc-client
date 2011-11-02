VERSION 5.00
Begin VB.Form frmPerform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit peform"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPerform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_labelManager As New CLabelManager
Private m_textPerform As VB.TextBox

Private Sub Form_Load()
    initControls
    updateColours Controls
End Sub

Private Sub initControls()
    Set m_textPerform = createControl(Controls, "VB.TextBox", "textPerform")
    m_textPerform.Move 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 50
End Sub

Private Sub Form_Paint()
    SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    
    SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    m_labelManager.renderLabels Me.hdc
End Sub
