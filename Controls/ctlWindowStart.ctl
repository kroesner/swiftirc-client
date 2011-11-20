VERSION 5.00
Begin VB.UserControl ctlWindowStart 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlWindowStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements ITabWindow
Implements IColourUser

Private m_fontManager As CFontManager

Private m_realWindow As VBControlExtender
Private m_tab As CTab

Private m_tabIndex As Long
Private WithEvents m_ctlStartPanel As swiftIrc.ctlStartPanel
Attribute m_ctlStartPanel.VB_VarHelpID = -1

Private m_client As swiftIrc.SwiftIrcClient

Public Event newSession()
Public Event profileConnect(profile As CServerProfile)

Public Property Get client() As swiftIrc.SwiftIrcClient
10        Set client = m_client
End Property

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
10        Set m_client = newValue
20        m_ctlStartPanel.client = m_client
End Property

Public Property Let switchbartab(newValue As CTab)
10        Set m_tab = newValue
End Property

Private Sub IColourUser_coloursUpdated()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
20        updateColours Controls
End Sub

Private Property Get ITabWindow_getTab() As CTab
10        Set ITabWindow_getTab = m_tab
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_ctlStartPanel_profileConnect(profile As CServerProfile)
10        RaiseEvent profileConnect(profile)
End Sub

Private Sub m_ctlStartPanel_newSession()
10        RaiseEvent newSession
End Sub

Private Sub UserControl_Initialize()
10        initControls
20        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
End Sub

Private Sub initControls()
10        Set m_ctlStartPanel = createControl(Controls, "swiftIrc.ctlStartPanel", "startPanel")
20        getRealWindow(m_ctlStartPanel).Move 0, 0, 700, 120
End Sub

Private Sub UserControl_Resize()
          Dim x As Long
          Dim y As Long
          
10        x = UserControl.ScaleWidth / 2 - (getRealWindow(m_ctlStartPanel).width / 2)
20        y = UserControl.ScaleHeight / 2 - (getRealWindow(m_ctlStartPanel).height / 2)
          
30        If x < 0 Then x = 0
40        If y < 0 Then y = 0
          
50        getRealWindow(m_ctlStartPanel).Move x, y, getRealWindow(m_ctlStartPanel).width, _
              getRealWindow(m_ctlStartPanel).height
End Sub
