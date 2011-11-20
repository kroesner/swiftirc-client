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
10        success = m_success
End Property

Public Property Get serverName() As String
10        serverName = m_fieldName.value
End Property

Public Property Let serverName(newValue As String)
10        m_fieldName.value = newValue
End Property

Public Property Get serverHost() As String
10        serverHost = m_fieldServerHost.value
End Property

Public Property Let serverHost(newValue As String)
10        m_fieldServerHost.value = newValue
End Property

Public Property Get serverPort() As Long
10        serverPort = Val(m_fieldServerPort.value)
End Property

Public Property Let serverPort(newValue As Long)
10        m_fieldServerPort.value = CStr(newValue)
End Property

Public Property Get serverPass() As String
10        serverPass = m_fieldServerPass.value
End Property

Public Property Let serverPass(newValue As String)
10        m_fieldServerPass.value = newValue
End Property

Public Property Get nickname() As String
10        nickname = m_fieldNickname.value
End Property

Public Property Let nickname(newValue As String)
10        m_fieldNickname.value = newValue
End Property

Public Property Get backupNickname() As String
10        backupNickname = m_fieldBackupNickname.value
End Property

Public Property Let backupNickname(newValue As String)
10        m_fieldBackupNickname.value = newValue
End Property

Public Property Get nicknamePassword() As String
10        nicknamePassword = m_fieldNicknamePassword.value
End Property

Public Property Let nicknamePassword(newValue As String)
10        m_fieldNicknamePassword.value = newValue
End Property

Public Property Get realName() As String
10        realName = m_fieldRealname.value
End Property

Public Property Let realName(newValue As String)
10        m_fieldRealname.value = newValue
End Property

Public Property Let editProfile(newValue As CServerProfile)
10        Set m_editProfile = newValue
End Property

Private Sub Form_Load()
10        initControls
20        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
30        updateColours Controls
          
40        If Not m_editProfile Is Nothing Then
50            m_fieldName.value = m_editProfile.name
60            m_fieldServerHost.value = m_editProfile.hostname
70            m_fieldServerPort.value = m_editProfile.port
80            m_fieldServerPass.value = m_editProfile.serverPassword
90            m_fieldNickname.value = m_editProfile.primaryNickname
100           m_fieldBackupNickname.value = m_editProfile.backupNickname
110           m_fieldRealname.value = m_editProfile.realName
120           m_fieldNicknamePassword.value = m_editProfile.nicknamePassword
              
130           m_enablePerform = m_editProfile.enablePerform
140           m_perform = m_editProfile.perform
              
150           m_checkAutoJoin.value = -m_editProfile.enableAutoJoin
160           m_checkAutoIdentify.value = -m_editProfile.enableAutoIdentify
170           m_checkReconnect.value = -m_editProfile.enableReconnect
180           m_checkConnectRetry.value = -m_editProfile.enableConnectRetry
              
190           updateAjList
200       End If
End Sub

Private Sub initControls()
10        m_labelManager.addLabel "Server details", ltHeading, 25, 15

20        Set m_fieldName = addField(Controls, "Name:", 25, 40, 300, 20)
30        Set m_fieldServerHost = addField(Controls, "Server hostname:", 25, 65, 300, 20)
40        Set m_fieldServerPort = addField(Controls, "Server port:", 25, 90, 200, 20)
50        Set m_fieldServerPass = addField(Controls, "Server password:", 25, 115, 300, 20)
          
60        m_labelManager.addLabel "Login details", ltHeading, 25, 145
          
70        Set m_fieldNickname = addField(Controls, "Nickname:", 25, 170, 300, 20)
80        Set m_fieldBackupNickname = addField(Controls, "Backup nickname:", 25, 195, 300, 20)
90        Set m_fieldNicknamePassword = addField(Controls, "Nickname password:", 25, 220, 300, 20)
100       Set m_fieldRealname = addField(Controls, "Real name:", 25, 245, 300, 20)
          
110       m_labelManager.addLabel "Auto join channels", ltHeading, 25, 275
          
120       Set m_listChannels = createControl(Controls, "VB.ListBox", "list")
130       m_listChannels.Move 25, 300, Me.ScaleWidth - 140, 85
          
140       Set m_buttonAddChannel = addButton(Controls, "&Add", 250, 300, 75, 20)
150       Set m_buttonEditChannel = addButton(Controls, "&Edit", 250, 325, 75, 20)
160       Set m_buttonRemoveChannel = addButton(Controls, "&Remove", 250, 350, 75, 20)
          
170       m_labelManager.addLabel "Perform", ltHeading, 25, 395
          
180       Set m_buttonEditPerform = addButton(Controls, "Edit perform", 25, 415, 90, 20)
          
190       m_labelManager.addLabel "Options", ltHeading, 25, 440
          
200       Set m_checkAutoJoin = addCheckBox(Controls, "Enable auto join channels", 25, 460, 200, 15)
210       Set m_checkAutoIdentify = addCheckBox(Controls, "Auto-identify with NickServ (where available)", _
              25, 480, 300, 15)
          
220       Set m_checkReconnect = addCheckBox(Controls, "Reconnect on disconnection", 25, 500, 200, 15)
230       Set m_checkConnectRetry = addCheckBox(Controls, "Auto-Retry connecting if unsuccessful", 25, _
              520, 300, 15)
          
240       Set m_buttonSave = addButton(Controls, "&Save", Me.ScaleWidth - 225, Me.ScaleHeight - 40, 100, _
              20)
250       Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 120, Me.ScaleHeight - 40, _
              100, 20)
          
260       m_fieldName.required = True
270       m_fieldNickname.required = True
280       m_fieldNickname.mask = fmIrcNickname
290       m_fieldBackupNickname.mask = fmIrcNickname
300       m_fieldServerHost.required = True
310       m_fieldServerPort.setFieldWidth 150, 50
          
320       m_fieldNicknamePassword.password = True
End Sub

Private Sub Form_Paint()
10        SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
          
30        SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
40        m_labelManager.renderLabels Me.hdc
End Sub

Private Sub m_buttonAddChannel_clicked()
          Dim channelEdit As New frmEditAutoJoinChannel
          
10        channelEdit.Show vbModal, Me
          
20        If channelEdit.success Then
30            If Not m_editProfile Is Nothing Then
40                m_editProfile.addAutoJoinChannel channelEdit.channel, channelEdit.key
50            Else
                  Dim newChannel As New CAutoJoinChannel
                  
60                newChannel.channel = channelEdit.channel
70                newChannel.key = channelEdit.key
                  
80                m_autoJoinChannels.Add newChannel
90            End If
              
100           updateAjList
110       End If
End Sub

Private Sub updateAjList()
          Dim count As Long
          
10        m_listChannels.clear
          
20        If Not m_editProfile Is Nothing Then
30            For count = 1 To m_editProfile.autoJoinChannelCount
40                If LenB(m_editProfile.autoJoinChannel(count).key) <> 0 Then
50                    m_listChannels.addItem m_editProfile.autoJoinChannel(count).channel & " (has key)"
60                Else
70                    m_listChannels.addItem m_editProfile.autoJoinChannel(count).channel
80                End If
90            Next count
100       Else
110           For count = 1 To m_autoJoinChannels.count
120               If LenB(m_autoJoinChannels.item(count).key) <> 0 Then
130                   m_listChannels.addItem m_autoJoinChannels.item(count).channel & " (has key)"
140               Else
150                   m_listChannels.addItem m_autoJoinChannels.item(count).channel
160               End If
170           Next count
180       End If
End Sub

Private Sub m_buttonEditChannel_clicked()
10        If m_listChannels.ListIndex <> -1 Then
              Dim channelEdit As New frmEditAutoJoinChannel
              
20            If Not m_editProfile Is Nothing Then
30                channelEdit.channel = m_editProfile.autoJoinChannel(m_listChannels.ListIndex + _
                      1).channel
40                channelEdit.key = m_editProfile.autoJoinChannel(m_listChannels.ListIndex + 1).key
50            Else
60                channelEdit.channel = m_autoJoinChannels.item(m_listChannels.ListIndex + 1).channel
70                channelEdit.key = m_autoJoinChannels.item(m_listChannels.ListIndex + 1).key
80            End If
              
90            channelEdit.Show vbModal, Me
              
100           If channelEdit.success Then
110               If Not m_editProfile Is Nothing Then
120                   m_editProfile.autoJoinChannel(m_listChannels.ListIndex + 1).channel = _
                          channelEdit.channel
130                   m_editProfile.autoJoinChannel(m_listChannels.ListIndex + 1).key = channelEdit.key
140               Else
150                   m_autoJoinChannels.item(m_listChannels.ListIndex + 1).channel = channelEdit.channel
160                   m_autoJoinChannels.item(m_listChannels.ListIndex + 1).key = channelEdit.key
170               End If
                  
180               updateAjList
190           End If
200       End If
End Sub

Private Sub m_buttonEditPerform_clicked()
          Dim perform As New frmPerform
          
10        perform.enablePerform = m_enablePerform
20        perform.perform = m_perform
          
30        perform.Show vbModal, Me
          
40        If Not perform.cancelled Then
50            m_enablePerform = perform.enablePerform
60            m_perform = perform.perform
70        End If
End Sub

Private Sub m_buttonRemoveChannel_clicked()
10        If m_listChannels.ListIndex <> -1 Then
20            If Not m_editProfile Is Nothing Then
30                m_editProfile.removeAutoJoinChannel m_listChannels.ListIndex + 1
40            Else
50                m_autoJoinChannels.Remove m_listChannels.ListIndex + 1
60            End If
70            updateAjList
80        End If
End Sub

Private Sub updateProfile(profile As CServerProfile)
10        profile.name = m_fieldName.value
20        profile.hostname = m_fieldServerHost.value
          
30        If Val(m_fieldServerPort.value) = 0 Then
40            profile.port = 6667
50        Else
60            profile.port = Val(m_fieldServerPort.value)
70        End If
          
80        profile.serverPassword = m_fieldServerPass.value
90        profile.primaryNickname = m_fieldNickname.value
100       profile.backupNickname = m_fieldBackupNickname.value
110       profile.realName = m_fieldRealname.value
120       profile.nicknamePassword = m_fieldNicknamePassword.value
          
130       profile.enablePerform = m_enablePerform
140       profile.perform = m_perform
          
150       profile.enableAutoJoin = -m_checkAutoJoin.value
160       profile.enableAutoIdentify = -m_checkAutoIdentify.value
170       profile.enableReconnect = -m_checkReconnect.value
180       profile.enableConnectRetry = -m_checkConnectRetry.value
End Sub

Private Sub m_buttonSave_clicked()
10        If LenB(m_fieldName.value) = 0 Then
20            MsgBox "You must provide a name for this server profile", vbCritical, "Missing information"
30            Exit Sub
40        End If
          
50        If LenB(m_fieldServerHost.value) = 0 Then
60            MsgBox "You must provide a server address", vbCritical, "Missing information"
70            Exit Sub
80        End If
          
          Dim profile As CServerProfile

90        If m_editProfile Is Nothing Then
100           If Not serverProfiles.findProfile(m_fieldName.value) Is Nothing Then
110               MsgBox "A profile with the name """ & m_fieldName.value & """ already exists", _
                      vbCritical, "Name already in use"
120               Exit Sub
130           End If
          
140           Set profile = New CServerProfile
              
150           updateProfile profile
              
              Dim count As Long
              
160           For count = 1 To m_autoJoinChannels.count
170               profile.addAutoJoinChannel m_autoJoinChannels.item(count).channel, _
                      m_autoJoinChannels.item(count).key
180           Next count
              
190           serverProfiles.addProfile profile
200       Else
210           If Not serverProfiles.findProfile(m_fieldName.value) Is Nothing And Not _
                  serverProfiles.findProfile(m_fieldName.value) Is m_editProfile Then
                  
220               MsgBox "A profile with the name """ & m_fieldName.value & """ already exists", _
                      vbCritical, "Name already in use"
230               Exit Sub
240           End If
          
250           If StrComp(m_fieldName.value, m_editProfile.name, vbTextCompare) = 0 Then
260               updateProfile m_editProfile
270           Else
280               serverProfiles.removeProfile m_editProfile.name
                  
290               Set profile = New CServerProfile
300               updateProfile profile
310               serverProfiles.addProfile profile
320           End If
330       End If

340       m_success = True
350       Me.Hide
End Sub

Private Sub m_buttonCancel_clicked()
10        Me.Hide
End Sub
