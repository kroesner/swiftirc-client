VERSION 5.00
Begin VB.Form frmEditServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit server"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "frmEditServer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   582
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_labelManager As New CLabelManager

Private m_fieldName As ctlField
Private m_fieldServerHost As ctlField
Private m_fieldServerPort As ctlField
Private m_fieldServerPass As ctlField

Private m_fieldNickname As ctlField
Private m_fieldBackupNickname As ctlField
Private m_fieldNicknamePassword As ctlField
Private m_fieldRealname As ctlField

Private m_listChannels As VB.ListBox
Private WithEvents m_buttonAddChannel As ctlButton
Attribute m_buttonAddChannel.VB_VarHelpID = -1
Private WithEvents m_buttonEditChannel As ctlButton
Attribute m_buttonEditChannel.VB_VarHelpID = -1
Private WithEvents m_buttonRemoveChannel As ctlButton
Attribute m_buttonRemoveChannel.VB_VarHelpID = -1

Private WithEvents m_buttonEditPerform As ctlButton
Attribute m_buttonEditPerform.VB_VarHelpID = -1

Private m_checkAutoJoin As VB.CheckBox
Private m_checkAutoIdentify As VB.CheckBox
Private m_checkReconnect As VB.CheckBox
Private m_checkConnectRetry As VB.CheckBox

Private WithEvents m_buttonSave As ctlButton
Attribute m_buttonSave.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1

Private m_enablePerform As Boolean
Private m_perform As String

Private m_success As Boolean
Private m_editProfile As CServerProfile
Private m_autoJoinChannels As New cArrayList

Public Property Get success() As Boolean
    success = m_success
End Property

Public Property Get serverName() As String
    serverName = m_fieldName.value
End Property

Public Property Let serverName(newValue As String)
    m_fieldName.value = newValue
End Property

Public Property Get serverHost() As String
    serverHost = m_fieldServerHost.value
End Property

Public Property Let serverHost(newValue As String)
    m_fieldServerHost.value = newValue
End Property

Public Property Get serverPort() As Long
    serverPort = Val(m_fieldServerPort.value)
End Property

Public Property Let serverPort(newValue As Long)
    m_fieldServerPort.value = CStr(newValue)
End Property

Public Property Get serverPass() As String
    serverPass = m_fieldServerPass.value
End Property

Public Property Let serverPass(newValue As String)
    m_fieldServerPass.value = newValue
End Property

Public Property Get nickname() As String
    nickname = m_fieldNickname.value
End Property

Public Property Let nickname(newValue As String)
    m_fieldNickname.value = newValue
End Property

Public Property Get backupNickname() As String
    backupNickname = m_fieldBackupNickname.value
End Property

Public Property Let backupNickname(newValue As String)
    m_fieldBackupNickname.value = newValue
End Property

Public Property Get nicknamePassword() As String
    nicknamePassword = m_fieldNicknamePassword.value
End Property

Public Property Let nicknamePassword(newValue As String)
    m_fieldNicknamePassword.value = newValue
End Property

Public Property Get realName() As String
    realName = m_fieldRealname.value
End Property

Public Property Let realName(newValue As String)
    m_fieldRealname.value = newValue
End Property

Public Property Let editProfile(newValue As CServerProfile)
    Set m_editProfile = newValue
End Property

Private Sub Form_Load()
    initControls
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    updateColours Controls
    
    If Not m_editProfile Is Nothing Then
        m_fieldName.value = m_editProfile.name
        m_fieldServerHost.value = m_editProfile.hostname
        m_fieldServerPort.value = m_editProfile.port
        m_fieldServerPass.value = m_editProfile.serverPassword
        m_fieldNickname.value = m_editProfile.primaryNickname
        m_fieldBackupNickname.value = m_editProfile.backupNickname
        m_fieldRealname.value = m_editProfile.realName
        m_fieldNicknamePassword.value = m_editProfile.nicknamePassword
        
        m_enablePerform = m_editProfile.enablePerform
        m_perform = m_editProfile.perform
        
        m_checkAutoJoin.value = -m_editProfile.enableAutoJoin
        m_checkAutoIdentify.value = -m_editProfile.enableAutoIdentify
        m_checkReconnect.value = -m_editProfile.enableReconnect
        m_checkConnectRetry.value = -m_editProfile.enableConnectRetry
        
        updateAjList
    End If
End Sub

Private Sub initControls()
    m_labelManager.addLabel "Server details", ltHeading, 25, 15

    Set m_fieldName = addField(Controls, "Name:", 25, 40, 300, 20)
    Set m_fieldServerHost = addField(Controls, "Server hostname:", 25, 65, 300, 20)
    Set m_fieldServerPort = addField(Controls, "Server port:", 25, 90, 200, 20)
    Set m_fieldServerPass = addField(Controls, "Server password:", 25, 115, 300, 20)
    
    m_labelManager.addLabel "Login details", ltHeading, 25, 145
    
    Set m_fieldNickname = addField(Controls, "Nickname:", 25, 170, 300, 20)
    Set m_fieldBackupNickname = addField(Controls, "Backup nickname:", 25, 195, 300, 20)
    Set m_fieldNicknamePassword = addField(Controls, "Nickname password:", 25, 220, 300, 20)
    Set m_fieldRealname = addField(Controls, "Real name:", 25, 245, 300, 20)
    
    m_labelManager.addLabel "Auto join channels", ltHeading, 25, 275
    
    Set m_listChannels = createControl(Controls, "VB.ListBox", "list")
    m_listChannels.Move 25, 300, Me.ScaleWidth - 140, 85
    
    Set m_buttonAddChannel = addButton(Controls, "&Add", 250, 300, 75, 20)
    Set m_buttonEditChannel = addButton(Controls, "&Edit", 250, 325, 75, 20)
    Set m_buttonRemoveChannel = addButton(Controls, "&Remove", 250, 350, 75, 20)
    
    m_labelManager.addLabel "Perform", ltHeading, 25, 395
    
    Set m_buttonEditPerform = addButton(Controls, "Edit perform", 25, 415, 90, 20)
    
    m_labelManager.addLabel "Options", ltHeading, 25, 440
    
    Set m_checkAutoJoin = addCheckBox(Controls, "Enable auto join channels", 25, 460, 200, 15)
    Set m_checkAutoIdentify = addCheckBox(Controls, "Auto-identify with NickServ (where available)", _
        25, 480, 300, 15)
    
    Set m_checkReconnect = addCheckBox(Controls, "Reconnect on disconnection", 25, 500, 200, 15)
    Set m_checkConnectRetry = addCheckBox(Controls, "Auto-Retry connecting if unsuccessful", 25, _
        520, 300, 15)
    
    Set m_buttonSave = addButton(Controls, "&Save", Me.ScaleWidth - 225, Me.ScaleHeight - 40, 100, _
        20)
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 120, Me.ScaleHeight - 40, _
        100, 20)
    
    m_fieldName.required = True
    m_fieldNickname.required = True
    m_fieldNickname.mask = fmIrcNickname
    m_fieldBackupNickname.mask = fmIrcNickname
    m_fieldServerHost.required = True
    m_fieldServerPort.setFieldWidth 150, 50
    
    m_fieldNicknamePassword.password = True
End Sub

Private Sub Form_Paint()
    SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    
    SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    m_labelManager.renderLabels Me.hdc
End Sub

Private Sub m_buttonAddChannel_clicked()
    Dim channelEdit As New frmEditAutoJoinChannel
    
    channelEdit.Show vbModal, Me
    
    If channelEdit.success Then
        If Not m_editProfile Is Nothing Then
            m_editProfile.addAutoJoinChannel channelEdit.channel, channelEdit.key
        Else
            Dim newChannel As New CAutoJoinChannel
            
            newChannel.channel = channelEdit.channel
            newChannel.key = channelEdit.key
            
            m_autoJoinChannels.Add newChannel
        End If
        
        updateAjList
    End If
End Sub

Private Sub updateAjList()
    Dim count As Long
    
    m_listChannels.clear
    
    If Not m_editProfile Is Nothing Then
        For count = 1 To m_editProfile.autoJoinChannelCount
            If LenB(m_editProfile.autoJoinChannel(count).key) <> 0 Then
                m_listChannels.addItem m_editProfile.autoJoinChannel(count).channel & " (has key)"
            Else
                m_listChannels.addItem m_editProfile.autoJoinChannel(count).channel
            End If
        Next count
    Else
        For count = 1 To m_autoJoinChannels.count
            If LenB(m_autoJoinChannels.item(count).key) <> 0 Then
                m_listChannels.addItem m_autoJoinChannels.item(count).channel & " (has key)"
            Else
                m_listChannels.addItem m_autoJoinChannels.item(count).channel
            End If
        Next count
    End If
End Sub

Private Sub m_buttonEditChannel_clicked()
    If m_listChannels.ListIndex <> -1 Then
        Dim channelEdit As New frmEditAutoJoinChannel
        
        If Not m_editProfile Is Nothing Then
            channelEdit.channel = m_editProfile.autoJoinChannel(m_listChannels.ListIndex + _
                1).channel
            channelEdit.key = m_editProfile.autoJoinChannel(m_listChannels.ListIndex + 1).key
        Else
            channelEdit.channel = m_autoJoinChannels.item(m_listChannels.ListIndex + 1).channel
            channelEdit.key = m_autoJoinChannels.item(m_listChannels.ListIndex + 1).key
        End If
        
        channelEdit.Show vbModal, Me
        
        If channelEdit.success Then
            If Not m_editProfile Is Nothing Then
                m_editProfile.autoJoinChannel(m_listChannels.ListIndex + 1).channel = _
                    channelEdit.channel
                m_editProfile.autoJoinChannel(m_listChannels.ListIndex + 1).key = channelEdit.key
            Else
                m_autoJoinChannels.item(m_listChannels.ListIndex + 1).channel = channelEdit.channel
                m_autoJoinChannels.item(m_listChannels.ListIndex + 1).key = channelEdit.key
            End If
            
            updateAjList
        End If
    End If
End Sub

Private Sub m_buttonEditPerform_clicked()
    Dim perform As New frmPerform
    
    perform.enablePerform = m_enablePerform
    perform.perform = m_perform
    
    perform.Show vbModal, Me
    
    If Not perform.cancelled Then
        m_enablePerform = perform.enablePerform
        m_perform = perform.perform
    End If
End Sub

Private Sub m_buttonRemoveChannel_clicked()
    If m_listChannels.ListIndex <> -1 Then
        If Not m_editProfile Is Nothing Then
            m_editProfile.removeAutoJoinChannel m_listChannels.ListIndex + 1
        Else
            m_autoJoinChannels.Remove m_listChannels.ListIndex + 1
        End If
        updateAjList
    End If
End Sub

Private Sub updateProfile(profile As CServerProfile)
    profile.name = m_fieldName.value
    profile.hostname = m_fieldServerHost.value
    
    If Val(m_fieldServerPort.value) = 0 Then
        profile.port = 6667
    Else
        profile.port = Val(m_fieldServerPort.value)
    End If
    
    profile.serverPassword = m_fieldServerPass.value
    profile.primaryNickname = m_fieldNickname.value
    profile.backupNickname = m_fieldBackupNickname.value
    profile.realName = m_fieldRealname.value
    profile.nicknamePassword = m_fieldNicknamePassword.value
    
    profile.enablePerform = m_enablePerform
    profile.perform = m_perform
    
    profile.enableAutoJoin = -m_checkAutoJoin.value
    profile.enableAutoIdentify = -m_checkAutoIdentify.value
    profile.enableReconnect = -m_checkReconnect.value
    profile.enableConnectRetry = -m_checkConnectRetry.value
End Sub

Private Sub m_buttonSave_clicked()
    If LenB(m_fieldName.value) = 0 Then
        MsgBox "You must provide a name for this server profile", vbCritical, "Missing information"
        Exit Sub
    End If
    
    If LenB(m_fieldServerHost.value) = 0 Then
        MsgBox "You must provide a server address", vbCritical, "Missing information"
        Exit Sub
    End If
    
    Dim profile As CServerProfile

    If m_editProfile Is Nothing Then
        If Not serverProfiles.findProfile(m_fieldName.value) Is Nothing Then
            MsgBox "A profile with the name """ & m_fieldName.value & """ already exists", _
                vbCritical, "Name already in use"
            Exit Sub
        End If
    
        Set profile = New CServerProfile
        
        updateProfile profile
        
        Dim count As Long
        
        For count = 1 To m_autoJoinChannels.count
            profile.addAutoJoinChannel m_autoJoinChannels.item(count).channel, _
                m_autoJoinChannels.item(count).key
        Next count
        
        serverProfiles.addProfile profile
    Else
        If Not serverProfiles.findProfile(m_fieldName.value) Is Nothing And Not _
            serverProfiles.findProfile(m_fieldName.value) Is m_editProfile Then
            
            MsgBox "A profile with the name """ & m_fieldName.value & """ already exists", _
                vbCritical, "Name already in use"
            Exit Sub
        End If
    
        If StrComp(m_fieldName.value, m_editProfile.name, vbTextCompare) = 0 Then
            updateProfile m_editProfile
        Else
            serverProfiles.removeProfile m_editProfile.name
            
            Set profile = New CServerProfile
            updateProfile profile
            serverProfiles.addProfile profile
        End If
    End If

    m_success = True
    Me.Hide
End Sub

Private Sub m_buttonCancel_clicked()
    Me.Hide
End Sub
