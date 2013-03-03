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
    address = m_fieldAddress.value
End Property

Public Property Let address(newValue As String)
    m_fieldAddress.value = newValue
End Property

Public Property Get success() As Boolean
    success = m_success
End Property

Private Sub Form_Initialize()
    initControls
End Sub

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    updateColours Controls
End Sub

Private Sub initControls()
    Set m_fieldAddress = addField(Controls, "Edit address:", 20, 20, Me.ScaleWidth - 40, 20)
    m_fieldAddress.setFieldWidth 100, 195
    Set m_buttonOk = addButton(Controls, "Ok", Me.ScaleWidth - 175, Me.ScaleHeight - 40, 75, 20)
    Set m_buttonCancel = addButton(Controls, "Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, 20)
End Sub

Private Sub Form_Paint()
    SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
End Sub

Private Sub m_buttonCancel_clicked()
    Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
    If LenB(m_fieldAddress.value) = 0 Then
        MsgBox "Please enter an address", vbCritical, "Missing information"
        Exit Sub
    End If
    
    m_success = True
    Me.Hide
End Sub
