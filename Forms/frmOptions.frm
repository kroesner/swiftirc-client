VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat options"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9000
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_client As swiftIrc.SwiftIrcClient

Private m_children As New cArrayList

Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1
Private WithEvents m_buttonApply As swiftIrc.ctlButton
Attribute m_buttonApply.VB_VarHelpID = -1
Private WithEvents m_tabStrip As swiftIrc.ctlTabStrip
Attribute m_tabStrip.VB_VarHelpID = -1

Private m_tabConnection As CTabStripItem
Private m_tabColour As CTabStripItem
Private m_tabColour2 As CTabStripItem
Private m_tabIrc As CTabStripItem

Private m_wndConnection As swiftIrc.ctlOptionsConnection
Private m_wndColour As swiftIrc.ctlOptionsColour
Private m_wndColour2 As swiftIrc.ctlOptionsColour2
Private m_wndIrc As swiftIrc.ctlOptionsIrc

Friend Property Get parent() As SwiftIrcClient
    Set parent = m_client
End Property

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
    Set m_client = newValue
    m_wndConnection.client = m_client
End Property

Private Sub colourUpdate()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    updateColours Controls
End Sub

Private Sub initControls()
    Set m_buttonOk = addButton(Controls, "&Ok", Me.ScaleWidth - 255, Me.ScaleHeight - 25, 75, 20)
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 170, Me.ScaleHeight - 25, 75, 20)
    Set m_buttonApply = addButton(Controls, "&Apply", Me.ScaleWidth - 85, Me.ScaleHeight - 25, 75, 20)
    
    Set m_tabStrip = createControl(Controls, "swiftirc.ctltabstrip", "test")
    
    Set m_wndConnection = createControl(Controls, "swiftirc.ctlOptionsConnection", "connection")
    Set m_wndColour = createControl(Controls, "swiftIrc.ctlOptionsColour", "colour")
    Set m_wndColour2 = createControl(Controls, "swiftIrc.ctlOptionsColour2", "colour2")
    Set m_wndIrc = createControl(Controls, "swiftIrc.ctlOptionsIrc", "irc")
    
    getRealWindow(m_wndConnection).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
    getRealWindow(m_wndColour).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
    getRealWindow(m_wndColour2).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
    getRealWindow(m_wndIrc).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
    getRealWindow(m_tabStrip).Move 10, 0, Me.ScaleWidth - 175, 30
    
    Set m_tabConnection = m_tabStrip.addTab("Co&nnection")
    Set m_tabColour = m_tabStrip.addTab("A&ppearance")
    Set m_tabColour2 = m_tabStrip.addTab("App&earance 2")
    Set m_tabIrc = m_tabStrip.addTab("&Irc")
    
    m_tabStrip.selectTab m_tabConnection, False
End Sub

Private Sub Form_Initialize()
    initControls
    colourUpdate
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        cancel = 1
        closeOptionsDialog
    End If
End Sub

Private Sub m_buttonApply_clicked()
    saveAllSettings
End Sub

Private Sub m_buttonCancel_clicked()
    closeOptionsDialog
End Sub

Private Sub m_buttonOk_clicked()
    saveAllSettings
    closeOptionsDialog
End Sub

Private Sub m_tabStrip_tabSelected(selectedTab As CTabStripItem)
    If selectedTab Is m_tabConnection Then
        getRealWindow(m_wndConnection).visible = True
        getRealWindow(m_wndColour).visible = False
        getRealWindow(m_wndColour2).visible = False
        getRealWindow(m_wndIrc).visible = False
    ElseIf selectedTab Is m_tabColour Then
        getRealWindow(m_wndColour).visible = True
        getRealWindow(m_wndColour2).visible = False
        getRealWindow(m_wndConnection).visible = False
        getRealWindow(m_wndIrc).visible = False
    ElseIf selectedTab Is m_tabColour2 Then
        getRealWindow(m_wndColour2).visible = True
        getRealWindow(m_wndColour).visible = False
        getRealWindow(m_wndConnection).visible = False
        getRealWindow(m_wndIrc).visible = False
    ElseIf selectedTab Is m_tabIrc Then
        getRealWindow(m_wndColour).visible = False
        getRealWindow(m_wndColour2).visible = False
        getRealWindow(m_wndConnection).visible = False
        getRealWindow(m_wndIrc).visible = True
    End If
End Sub

Friend Sub saveSettings()
    m_wndColour.saveSettings
    m_wndColour2.saveSettings
    m_wndIrc.saveSettings
End Sub
