VERSION 5.00
Begin VB.UserControl ctlOptionsConnection 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlOptionsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IColourUser

Private m_client As swiftIrc.SwiftIrcClient

Private m_labelManager As New CLabelManager
Private m_realWindow As VBControlExtender

Private WithEvents m_listServerProfiles As VB.ListBox
Attribute m_listServerProfiles.VB_VarHelpID = -1
Private WithEvents m_buttonConnect As swiftIrc.ctlButton
Attribute m_buttonConnect.VB_VarHelpID = -1
Private m_checkNewWindow As VB.CheckBox

Private WithEvents m_buttonAddServer As swiftIrc.ctlButton
Attribute m_buttonAddServer.VB_VarHelpID = -1
Private WithEvents m_buttonEditServer As swiftIrc.ctlButton
Attribute m_buttonEditServer.VB_VarHelpID = -1
Private WithEvents m_buttonRemoveServer As swiftIrc.ctlButton
Attribute m_buttonRemoveServer.VB_VarHelpID = -1

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
    Set m_client = newValue
End Property

Private Sub IColourUser_coloursUpdated()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
    updateColours Controls
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_buttonAddServer_clicked()
    Dim editServer As New frmEditServer
    
    editServer.Show vbModal, UserControl.parent
    updateServerProfileList
    saveSettings
End Sub

Private Sub m_buttonConnect_clicked()
    If m_listServerProfiles.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim profile As CServerProfile
    Dim profileName As String
    
    Set profile = serverProfiles.profileItem(m_listServerProfiles.ListIndex + 1)
    
    If m_checkNewWindow.value = 1 Then
        Dim session As CSession
        Set session = m_client.newSession
        
        m_client.ShowWindow session.statusWindow
        
        session.serverProfile = profile
        session.connect
    End If
    
    UserControl.parent.Hide
End Sub

Private Sub m_buttonEditServer_clicked()
    If m_listServerProfiles.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim editServer As New frmEditServer
    
    editServer.editProfile = serverProfiles.profileItem(m_listServerProfiles.ListIndex + 1)
    editServer.Show vbModal, Me
    updateServerProfileList
    saveSettings
End Sub

Private Sub m_buttonRemoveServer_clicked()
    If m_listServerProfiles.ListIndex = -1 Then
        Exit Sub
    End If
    
    serverProfiles.removeProfileIndex m_listServerProfiles.ListIndex + 1
    updateServerProfileList
    saveSettings
End Sub

Private Sub UserControl_Initialize()
    initControls
    updateServerProfileList
End Sub

Private Sub initControls()
    m_labelManager.addLabel "Connect to a server", ltHeading, 15, 15
    Set m_listServerProfiles = createControl(Controls, "VB.ListBox", "list")
    m_listServerProfiles.Move 15, 35, 200, 285
    
    Set m_checkNewWindow = addCheckBox(Controls, "New window", 220, 35, 100, 20)
    m_checkNewWindow.value = 1
    
    Set m_buttonConnect = addButton(Controls, "Connect", 220, 60, 75, 25)
    
    Set m_buttonAddServer = addButton(Controls, "Add", 220, 110, 75, 25)
    Set m_buttonEditServer = addButton(Controls, "Edit", 220, 140, 75, 25)
    Set m_buttonRemoveServer = addButton(Controls, "Remove", 220, 170, 75, 25)
End Sub

Private Sub updateServerProfileList()
    Dim serverProfile As CServerProfile
    Dim count As Long
    
    m_listServerProfiles.clear
    
    For count = 1 To serverProfiles.profileCount
        Set serverProfile = serverProfiles.profileItem(count)
        m_listServerProfiles.addItem serverProfile.name & " (" & serverProfile.hostname & ":" & _
            serverProfile.port & ")"
    Next count
End Sub

Private Sub reDraw()
    m_labelManager.renderLabels (UserControl.hdc)
End Sub

Private Sub UserControl_Paint()
    reDraw
End Sub

Public Sub saveSettings()
    serverProfiles.saveProfiles g_userPath & "swiftirc_servers.xml"
End Sub

Private Sub UserControl_Terminate()
    debugLog "ctlOptionsConnection terminating"
End Sub
