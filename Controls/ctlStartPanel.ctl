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
    Set client = m_client
End Property

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
    Set m_client = newValue
End Property

Private Sub IColourUser_coloursUpdated()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    updateColours Controls
    UserControl_Paint
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_buttonConnect_clicked()
    settings.quickConnectNickname = m_fieldNickname.value
    settings.quickConnectChannel = m_fieldChannel.value
    settings.saveSettings
    
    If comboProfiles.ListIndex = 0 Then
        If LenB(settings.quickConnectNickname) = 0 Then
            serverProfiles.quickConnectProfile.primaryNickname = m_client.generateDefaultNickname
            serverProfiles.quickConnectProfile.backupNickname = m_client.generateDefaultNickname
            serverProfiles.quickConnectProfile.realName = m_client.generateDefaultNickname
        Else
            serverProfiles.quickConnectProfile.primaryNickname = settings.quickConnectNickname
            serverProfiles.quickConnectProfile.backupNickname = settings.quickConnectNickname & "_"
            serverProfiles.quickConnectProfile.realName = settings.quickConnectNickname
        End If
    
        If LenB(settings.quickConnectPassword) <> 0 Then
            serverProfiles.quickConnectProfile.enableAutoIdentify = True
            serverProfiles.quickConnectProfile.nicknamePassword = settings.quickConnectPassword
        Else
            serverProfiles.quickConnectProfile.enableAutoIdentify = False
        End If
        
        serverProfiles.quickConnectProfile.clearAutoJoinChannels
        
        If LenB(settings.quickConnectChannel) <> 0 Then
            serverProfiles.quickConnectProfile.enableAutoJoin = True
            serverProfiles.quickConnectProfile.addAutoJoinChannel settings.quickConnectChannel, vbNullString
        Else
            serverProfiles.quickConnectProfile.enableAutoJoin = False
        End If
        
        settings.setting("lastServerProfile", estString) = vbNullString
        RaiseEvent profileConnect(serverProfiles.quickConnectProfile)
    Else
        Dim profile As CServerProfile
        
        Set profile = serverProfiles.profileItem(comboProfiles.ListIndex)
        settings.setting("lastServerProfile", estString) = profile.name
        
        RaiseEvent profileConnect(profile)
    End If
End Sub

Private Sub m_buttonNewServerWindow_clicked()
    RaiseEvent newSession
End Sub

Private Sub m_buttonOptions_clicked()
    openOptionsDialog m_client
End Sub

Private Sub setProfile()
    If comboProfiles.ListIndex = 0 Then
        m_fieldNickname.visible = True
        m_fieldChannel.visible = True
        m_buttonRegister.visible = False
        m_buttonSetPassword.visible = True
        
        m_labelProfileName.visible = False
        m_labelProfileServer.visible = False
        m_labelProfileNickname.visible = False
    Else
        m_fieldNickname.visible = False
        m_fieldChannel.visible = False
        m_buttonRegister.visible = False
        m_buttonSetPassword.visible = False
        
        m_labelProfileName.caption = "Profile name: " & serverProfiles.profileItem(comboProfiles.ListIndex).name
        m_labelProfileServer.caption = "Server hostname: " & serverProfiles.profileItem(comboProfiles.ListIndex).hostname
        m_labelProfileNickname.caption = "Nickname: " & serverProfiles.profileItem(comboProfiles.ListIndex).primaryNickname
         
        m_labelProfileName.visible = True
        m_labelProfileServer.visible = True
        m_labelProfileNickname.visible = True
    End If
    
    UserControl_Paint
End Sub

Private Sub comboProfiles_Click()
    setProfile
End Sub

Private Sub m_buttonRegister_clicked()
    Dim register As New frmRegisterPhase1
    
    register.init m_client
    register.Show vbModal, Me
    Unload register
End Sub

Private Sub m_buttonSetPassword_clicked()
    Dim result As Variant
    
    result = requestInput("Auto-identify", "Enter your nickname password", vbNullString, Me, True)
    
    If result = False Or result = vbNullString Then
        Exit Sub
    End If
    
    settings.quickConnectPassword = result
    settings.saveSettings
End Sub

Private Sub m_serverProfiles_profilesChanged()
    If comboProfiles.ListIndex <> 0 Then
        Dim currentProfileName As String
        currentProfileName = comboProfiles.list(comboProfiles.ListIndex)
        
        updateProfiles
        
        Dim count As Long
        
        For count = 1 To comboProfiles.ListCount - 1
            If StrComp(comboProfiles.list(count), currentProfileName, vbTextCompare) = 0 Then
                comboProfiles.ListIndex = count
                setProfile
                Exit Sub
            End If
        Next count
        
        comboProfiles.ListIndex = 0
        setProfile
    Else
        updateProfiles
        comboProfiles.ListIndex = 0
        setProfile
    End If
End Sub

Private Sub UserControl_Initialize()
    initControls
    IColourUser_coloursUpdated
    
    Set m_serverProfiles = serverProfiles
    
    m_fieldNickname.value = settings.quickConnectNickname
    m_fieldChannel.value = settings.quickConnectChannel
End Sub

Private Sub updateProfiles()
    Dim count As Long
    
    comboProfiles.clear
    comboProfiles.addItem "Quick connect", 0
    
    For count = 1 To serverProfiles.profileCount
        comboProfiles.addItem serverProfiles.profileItem(count).name
    Next count
End Sub

Private Sub initControls()
    Set m_fieldNickname = addField(Controls, "Nickname:", 225, 35, 200, 20)
    Set m_fieldChannel = addField(Controls, "Channel:", 225, 60, 200, 20)
    
    m_fieldNickname.justification = fjRight
    m_fieldChannel.justification = fjRight
    
    m_fieldNickname.mask = fmIrcNickname
    
    Set m_buttonRegister = addButton(Controls, "Register nickname", 450, 35, 125, 20)
    
    m_buttonRegister.visible = False
    
    Set m_buttonSetPassword = addButton(Controls, "Set-up auto-identify", 450, 35, 125, 20)
    
    m_labelManager.addLabel "Select profile:", ltSubHeading, 225, 5
    
    Set m_labelProfileName = m_labelManager.addLabel("Name", ltNormal, 255, 35)
    Set m_labelProfileServer = m_labelManager.addLabel("Server", ltNormal, 255, 50)
    Set m_labelProfileNickname = m_labelManager.addLabel("Nick", ltNormal, 255, 65)
    
    comboProfiles.left = 325
    comboProfiles.top = 5
    comboProfiles.width = 150

    updateProfiles
    comboProfiles.ListIndex = 0
    
    Dim count As Long
    Dim lastProfile As String
    
    lastProfile = settings.setting("lastServerProfile", estString)
    
    If LenB(lastProfile) <> 0 Then
        For count = 1 To serverProfiles.profileCount
            If StrComp(serverProfiles.profileItem(count).name, lastProfile, vbTextCompare) = 0 Then
                comboProfiles.ListIndex = count
            End If
        Next count
    End If

    setProfile
    
    Set m_buttonConnect = addButton(Controls, "Connect", 315, 90, 75, 20)
    Set m_buttonOptions = addButton(Controls, "Options", 10, 10, 125, 20)
    Set m_buttonNewServerWindow = addButton(Controls, "New server window", 10, 35, 125, 20)
End Sub

Private Sub UserControl_Paint()
    FillRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    m_labelManager.renderLabels UserControl.hdc
End Sub
