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

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
10        Set m_client = newValue
20        m_wndConnection.client = m_client
30        m_wndColour.client = m_client
40        m_wndColour2.client = m_client
End Property

Private Sub colourUpdate()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
20        updateColours Controls
End Sub

Private Sub initControls()
10        Set m_buttonOk = addButton(Controls, "&Ok", Me.ScaleWidth - 255, Me.ScaleHeight - 25, 75, 20)
20        Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 170, Me.ScaleHeight - 25, 75, 20)
30        Set m_buttonApply = addButton(Controls, "&Apply", Me.ScaleWidth - 85, Me.ScaleHeight - 25, 75, 20)
          
40        Set m_tabStrip = createControl(Controls, "swiftirc.ctltabstrip", "test")
          
50        Set m_wndConnection = createControl(Controls, "swiftirc.ctlOptionsConnection", "connection")
60        Set m_wndColour = createControl(Controls, "swiftIrc.ctlOptionsColour", "colour")
70        Set m_wndColour2 = createControl(Controls, "swiftIrc.ctlOptionsColour2", "colour2")
80        Set m_wndIrc = createControl(Controls, "swiftIrc.ctlOptionsIrc", "irc")
          
90        getRealWindow(m_wndConnection).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
100       getRealWindow(m_wndColour).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
110       getRealWindow(m_wndColour2).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
120       getRealWindow(m_wndIrc).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 60
130       getRealWindow(m_tabStrip).Move 10, 0, Me.ScaleWidth - 175, 30
          
140       Set m_tabConnection = m_tabStrip.addTab("Co&nnection")
150       Set m_tabColour = m_tabStrip.addTab("A&ppearance")
160       Set m_tabColour2 = m_tabStrip.addTab("App&earance 2")
170       Set m_tabIrc = m_tabStrip.addTab("&Irc")
          
180       m_tabStrip.selectTab m_tabConnection, False
End Sub

Private Sub Form_Initialize()
10        initControls
20        colourUpdate
End Sub

Private Sub m_buttonApply_clicked()
10        applySettings
End Sub

Private Sub m_buttonCancel_clicked()
10        debugLog "Options cancelled"
20        Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
10        applySettings
20        Me.Hide
End Sub

Private Sub m_tabStrip_tabSelected(selectedTab As CTabStripItem)
10        If selectedTab Is m_tabConnection Then
20            getRealWindow(m_wndConnection).visible = True
30            getRealWindow(m_wndColour).visible = False
40            getRealWindow(m_wndColour2).visible = False
50            getRealWindow(m_wndIrc).visible = False
60        ElseIf selectedTab Is m_tabColour Then
70            getRealWindow(m_wndColour).visible = True
80            getRealWindow(m_wndColour2).visible = False
90            getRealWindow(m_wndConnection).visible = False
100           getRealWindow(m_wndIrc).visible = False
110       ElseIf selectedTab Is m_tabColour2 Then
120           getRealWindow(m_wndColour2).visible = True
130           getRealWindow(m_wndColour).visible = False
140           getRealWindow(m_wndConnection).visible = False
150           getRealWindow(m_wndIrc).visible = False
160       ElseIf selectedTab Is m_tabIrc Then
170           getRealWindow(m_wndColour).visible = False
180           getRealWindow(m_wndColour2).visible = False
190           getRealWindow(m_wndConnection).visible = False
200           getRealWindow(m_wndIrc).visible = True
210       End If
End Sub

Private Sub applySettings()
10        m_wndConnection.saveSettings
20        m_wndColour.saveSettings
30        m_wndColour2.saveSettings
40        m_wndIrc.saveSettings
          
50        settings.saveSettings
60        m_client.coloursUpdated
End Sub
