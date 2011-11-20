VERSION 5.00
Begin VB.UserControl ctlStartPanel 
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   Begin VB.ComboBox comboProfiles 
      Height          =   315
      Left            =   465
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   1905
   End
End
Attribute VB_Name = "ctlStartPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IColourUser


Public Event newSession()
Public Event profileConnect(profile As CServerProfile)

Private m_client As swiftIrc.SwiftIrcClient

Private m_realWindow As VBControlExtender
Private m_labelManager As New CLabelManager
Private m_fieldNickname As swiftIrc.ctlField
Private m_fieldChannel As swiftIrc.ctlField

Private m_quickConnectPassword As String

Private m_labelProfileName As CLabel
Private m_labelProfileServer As CLabel
Private m_labelProfilePort As CLabel
Private m_labelProfileNickname As CLabel

Private WithEvents m_serverProfiles As CServerProfileManager
Attribute m_serverProfiles.VB_VarHelpID = -1

Private WithEvents m_buttonConnect As swiftIrc.ctlButton
Attribute m_buttonConnect.VB_VarHelpID = -1

Private WithEvents m_buttonOptions As swiftIrc.ctlButton
Attribute m_buttonOptions.VB_VarHelpID = -1

Private WithEvents m_buttonRegister As swiftIrc.ctlButton
Attribute m_buttonRegister.VB_VarHelpID = -1
Private WithEvents m_buttonSetPassword As swiftIrc.ctlButton
Attribute m_buttonSetPassword.VB_VarHelpID = -1

Private WithEvents m_buttonNewServerWindow As swiftIrc.ctlButton
Attribute m_buttonNewServerWindow.VB_VarHelpID = -1

Public Property Get client() As swiftIrc.SwiftIrcClient
10        Set client = m_client
End Property

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
10        Set m_client = newValue
End Property

Private Sub IColourUser_coloursUpdated()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
20        updateColours Controls
30        UserControl_Paint
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_buttonConnect_clicked()
10        settings.quickConnectNickname = m_fieldNickname.value
20        settings.quickConnectChannel = m_fieldChannel.value
30        settings.saveSettings
          
40        If comboProfiles.ListIndex = 0 Then
50            If LenB(settings.quickConnectNickname) = 0 Then
60                serverProfiles.quickConnectProfile.primaryNickname = m_client.generateDefaultNickname
70                serverProfiles.quickConnectProfile.backupNickname = m_client.generateDefaultNickname
80                serverProfiles.quickConnectProfile.realName = m_client.generateDefaultNickname
90            Else
100               serverProfiles.quickConnectProfile.primaryNickname = settings.quickConnectNickname
110               serverProfiles.quickConnectProfile.backupNickname = settings.quickConnectNickname & "_"
120               serverProfiles.quickConnectProfile.realName = settings.quickConnectNickname
130           End If
          
140           If LenB(settings.quickConnectPassword) <> 0 Then
150               serverProfiles.quickConnectProfile.enableAutoIdentify = True
160               serverProfiles.quickConnectProfile.nicknamePassword = settings.quickConnectPassword
170           Else
180               serverProfiles.quickConnectProfile.enableAutoIdentify = False
190           End If
              
200           serverProfiles.quickConnectProfile.clearAutoJoinChannels
              
210           If LenB(settings.quickConnectChannel) <> 0 Then
220               serverProfiles.quickConnectProfile.enableAutoJoin = True
230               serverProfiles.quickConnectProfile.addAutoJoinChannel settings.quickConnectChannel, vbNullString
240           Else
250               serverProfiles.quickConnectProfile.enableAutoJoin = False
260           End If
              
270           settings.setting("lastServerProfile", estString) = vbNullString
280           RaiseEvent profileConnect(serverProfiles.quickConnectProfile)
290       Else
              Dim profile As CServerProfile
              
300           Set profile = serverProfiles.profileItem(comboProfiles.ListIndex)
310           settings.setting("lastServerProfile", estString) = profile.name
              
320           RaiseEvent profileConnect(profile)
330       End If
End Sub

Private Sub m_buttonNewServerWindow_clicked()
10        RaiseEvent newSession
End Sub

Private Sub m_buttonOptions_clicked()
          Dim options As New frmOptions
          
10        options.client = m_client
20        options.Show vbModal, Me
          
30        Unload options
End Sub

Private Sub setProfile()
10        If comboProfiles.ListIndex = 0 Then
20            m_fieldNickname.visible = True
30            m_fieldChannel.visible = True
40            m_buttonRegister.visible = False
50            m_buttonSetPassword.visible = True
              
60            m_labelProfileName.visible = False
70            m_labelProfileServer.visible = False
80            m_labelProfileNickname.visible = False
90        Else
100           m_fieldNickname.visible = False
110           m_fieldChannel.visible = False
120           m_buttonRegister.visible = False
130           m_buttonSetPassword.visible = False
              
140           m_labelProfileName.caption = "Profile name: " & serverProfiles.profileItem(comboProfiles.ListIndex).name
150           m_labelProfileServer.caption = "Server hostname: " & serverProfiles.profileItem(comboProfiles.ListIndex).hostname
160           m_labelProfileNickname.caption = "Nickname: " & serverProfiles.profileItem(comboProfiles.ListIndex).primaryNickname
               
170           m_labelProfileName.visible = True
180           m_labelProfileServer.visible = True
190           m_labelProfileNickname.visible = True
200       End If
          
210       UserControl_Paint
End Sub

Private Sub comboProfiles_Click()
10        setProfile
End Sub

Private Sub m_buttonRegister_clicked()
          Dim register As New frmRegisterPhase1
          
10        register.init m_client
20        register.Show vbModal, Me
30        Unload register
End Sub

Private Sub m_buttonSetPassword_clicked()
          Dim result As Variant
          
10        result = requestInput("Auto-identify", "Enter your nickname password", vbNullString, Me, True)
          
20        If result = False Or result = vbNullString Then
30            Exit Sub
40        End If
          
50        settings.quickConnectPassword = result
60        settings.saveSettings
End Sub

Private Sub m_serverProfiles_profilesChanged()
10        If comboProfiles.ListIndex <> 0 Then
              Dim currentProfileName As String
20            currentProfileName = comboProfiles.list(comboProfiles.ListIndex)
              
30            updateProfiles
              
              Dim count As Long
              
40            For count = 1 To comboProfiles.ListCount - 1
50                If StrComp(comboProfiles.list(count), currentProfileName, vbTextCompare) = 0 Then
60                    comboProfiles.ListIndex = count
70                    setProfile
80                    Exit Sub
90                End If
100           Next count
              
110           comboProfiles.ListIndex = 0
120           setProfile
130       Else
140           updateProfiles
150           comboProfiles.ListIndex = 0
160           setProfile
170       End If
End Sub

Private Sub UserControl_Initialize()
10        initControls
20        IColourUser_coloursUpdated
          
30        Set m_serverProfiles = serverProfiles
          
40        m_fieldNickname.value = settings.quickConnectNickname
50        m_fieldChannel.value = settings.quickConnectChannel
End Sub

Private Sub updateProfiles()
          Dim count As Long
          
10        comboProfiles.clear
20        comboProfiles.addItem "Quick connect", 0
          
30        For count = 1 To serverProfiles.profileCount
40            comboProfiles.addItem serverProfiles.profileItem(count).name
50        Next count
End Sub

Private Sub initControls()
10        Set m_fieldNickname = addField(Controls, "Nickname:", 225, 35, 200, 20)
20        Set m_fieldChannel = addField(Controls, "Channel:", 225, 60, 200, 20)
          
30        m_fieldNickname.justification = fjRight
40        m_fieldChannel.justification = fjRight
          
50        m_fieldNickname.mask = fmIrcNickname
          
60        Set m_buttonRegister = addButton(Controls, "Register nickname", 450, 35, 125, 20)
          
70        m_buttonRegister.visible = False
          
80        Set m_buttonSetPassword = addButton(Controls, "Set-up auto-identify", 450, 35, 125, 20)
          
90        m_labelManager.addLabel "Select profile:", ltSubHeading, 225, 5
          
100       Set m_labelProfileName = m_labelManager.addLabel("Name", ltNormal, 255, 35)
110       Set m_labelProfileServer = m_labelManager.addLabel("Server", ltNormal, 255, 50)
120       Set m_labelProfileNickname = m_labelManager.addLabel("Nick", ltNormal, 255, 65)
          
130       comboProfiles.left = 325
140       comboProfiles.top = 5
150       comboProfiles.width = 150

160       updateProfiles
170       comboProfiles.ListIndex = 0
          
          Dim count As Long
          Dim lastProfile As String
          
180       lastProfile = settings.setting("lastServerProfile", estString)
          
190       If LenB(lastProfile) <> 0 Then
200           For count = 1 To serverProfiles.profileCount
210               If StrComp(serverProfiles.profileItem(count).name, lastProfile, vbTextCompare) = 0 Then
220                   comboProfiles.ListIndex = count
230               End If
240           Next count
250       End If

260       setProfile
          
270       Set m_buttonConnect = addButton(Controls, "Connect", 315, 90, 75, 20)
280       Set m_buttonOptions = addButton(Controls, "Options", 10, 10, 125, 20)
290       Set m_buttonNewServerWindow = addButton(Controls, "New server window", 10, 35, 125, 20)
End Sub

Private Sub UserControl_Paint()
10        FillRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
20        m_labelManager.renderLabels UserControl.hdc
End Sub
