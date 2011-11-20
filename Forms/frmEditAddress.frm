VERSION 5.00
Begin VB.Form frmEditAddress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit address"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5040
   Icon            =   "frmEditAddress.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_fieldAddress As swiftIrc.ctlField
Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1

Private m_success As Boolean

Public Property Get address() As String
10        address = m_fieldAddress.value
End Property

Public Property Let address(newValue As String)
10        m_fieldAddress.value = newValue
End Property

Public Property Get success() As Boolean
10        success = m_success
End Property

Private Sub Form_Initialize()
10        initControls
End Sub

Private Sub Form_Load()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
20        updateColours Controls
End Sub

Private Sub initControls()
10        Set m_fieldAddress = addField(Controls, "Edit address:", 20, 20, Me.ScaleWidth - 40, 20)
20        m_fieldAddress.setFieldWidth 100, 195
30        Set m_buttonOk = addButton(Controls, "Ok", Me.ScaleWidth - 175, Me.ScaleHeight - 40, 75, 20)
40        Set m_buttonCancel = addButton(Controls, "Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, _
              20)
End Sub

Private Sub Form_Paint()
10        SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
End Sub

Private Sub m_buttonCancel_clicked()
10        Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
10        If LenB(m_fieldAddress.value) = 0 Then
20            MsgBox "Please enter an address", vbCritical, "Missing information"
30            Exit Sub
40        End If
          
50        m_success = True
60        Me.Hide
End Sub
