VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CServerProfileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_serverProfiles As New Collection
Private m_quickConnectProfile As CServerProfile

Private Const PROFILES_FILENAME As String = "swiftirc_servers.xml"

Public Event profilesChanged()

Public Property Get quickConnectProfile() As CServerProfile
    Set quickConnectProfile = m_quickConnectProfile
End Property

Public Sub addProfile(profile As CServerProfile)
    If Not findProfile(profile.name) Is Nothing Then
        Exit Sub
    End If
    
    m_serverProfiles.Add profile, LCase$(profile.name)
    RaiseEvent profilesChanged
End Sub

Public Sub removeProfileIndex(index As Long)
    m_serverProfiles.Remove index
    RaiseEvent profilesChanged
End Sub

Public Sub removeProfile(name As String)
    On Error Resume Next
    m_serverProfiles.Remove LCase$(name)
    RaiseEvent profilesChanged
End Sub

Public Function findProfile(name As String) As CServerProfile
    On Error Resume Next
    Set findProfile = m_serverProfiles.item(LCase$(name))
End Function

Public Function profileItem(index As Long) As CServerProfile
    Set profileItem = m_serverProfiles.item(index)
End Function

Public Function profileCount() As Long
    profileCount = m_serverProfiles.count
End Function

Public Sub loadProfiles()
    Dim xml As New DOMDocument30
    Dim profileNode As IXMLDOMNode
    Dim filename As String
    
    filename = combinePath(g_userPath, PROFILES_FILENAME)
    
    If Not xml.load(filename) Then
        Exit Sub
    End If
    
    Dim profile As CServerProfile
    Dim channelsNode As IXMLDOMNode
    Dim channelNode As IXMLDOMNode
    
    For Each profileNode In xml.documentElement.childNodes
        Set profile = New CServerProfile
        
        profile.name = profileNode.Attributes.getNamedItem("name").text
        profile.hostname = xmlGetElementText(profileNode, "hostname")
        profile.port = Val(xmlGetElementText(profileNode, "port"))
        profile.serverPassword = xmlGetElementText(profileNode, "serverPassword")
        profile.primaryNickname = xmlGetElementText(profileNode, "primaryNickname")
        profile.backupNickname = xmlGetElementText(profileNode, "backupNickname")
        profile.realName = xmlGetElementText(profileNode, "realName")
        
        If xmlGetAttributeText(profileNode, "nicknamePassword", "crypt") = "yes" Then
            profile.nicknamePassword = decrypt(g_cryptKey, xmlGetElementText(profileNode, "nicknamePassword"))
        Else
            profile.nicknamePassword = xmlGetElementText(profileNode, "nicknamePassword")
        End If
        
        profile.enablePerform = xmlElementExists(profileNode, "enablePerform")
        profile.perform = Replace(xmlGetElementText(profileNode, "perform"), "\n", vbCrLf)
        
        profile.enableAutoJoin = xmlElementExists(profileNode, "enableAutoJoin")
        profile.enableAutoIdentify = xmlElementExists(profileNode, "enableAutoIdentify")
        
        profile.enableReconnect = xmlElementExists(profileNode, "enableReconnect")
        profile.enableConnectRetry = xmlElementExists(profileNode, "enableConnectRetry")
        
        Set channelsNode = profileNode.selectSingleNode("autoJoinChannels")
        
        If Not channelsNode Is Nothing Then
            For Each channelNode In channelsNode.childNodes
                profile.addAutoJoinChannel xmlGetElementText(channelNode, "name"), _
                    xmlGetElementText(channelNode, "key")
            Next channelNode
        End If
        
        addProfile profile
    Next profileNode
    
    RaiseEvent profilesChanged
End Sub

Public Sub saveProfiles()
    Dim profile As CServerProfile
    Dim xml As New DOMDocument30
    Dim profileNode As IXMLDOMNode
    Dim node As IXMLDOMNode
    Dim attrib As IXMLDOMAttribute
    
    Dim filename As String
    
    filename = combinePath(g_userPath, PROFILES_FILENAME)
    
    Set xml.documentElement = xml.createNode(NODE_ELEMENT, "profiles", "")
    
    For Each profile In m_serverProfiles
        Set profileNode = xml.createNode(NODE_ELEMENT, "profile", "")
        Set attrib = xml.createNode(NODE_ATTRIBUTE, "name", "")
        
        attrib.text = profile.name
        profileNode.Attributes.setNamedItem attrib
        xml.documentElement.appendChild profileNode
        
        xmlAddElement xml, profileNode, "hostname", profile.hostname
        
        If profile.port <> 0 Then
            xmlAddElement xml, profileNode, "port", profile.port
        End If
        
        If LenB(profile.serverPassword) <> 0 Then
            xmlAddElement xml, profileNode, "serverPassword", profile.serverPassword
        End If
        
        xmlAddElement xml, profileNode, "primaryNickname", profile.primaryNickname
        
        If LenB(profile.backupNickname) <> 0 Then
            xmlAddElement xml, profileNode, "backupNickname", profile.backupNickname
        End If
        
        If LenB(profile.realName) <> 0 Then
            xmlAddElement xml, profileNode, "realName", profile.realName
        End If
        
        If LenB(profile.nicknamePassword) <> 0 Then
            Dim passNode As IXMLDOMNode
            Set passNode = xmlAddElement(xml, profileNode, "nicknamePassword", encrypt(g_cryptKey, profile.nicknamePassword))
            
            Set attrib = xml.createAttribute("crypt")
            attrib.text = "yes"
            passNode.Attributes.setNamedItem attrib
        End If
        
        If profile.enablePerform Then
            xmlAddElement xml, profileNode, "enablePerform", vbNullString
        End If
        
        If LenB(profile.perform) <> 0 Then
            xmlAddElement xml, profileNode, "perform", Replace(profile.perform, vbCrLf, "\n")
        End If
        
        If profile.enableAutoJoin Then
            xmlAddElement xml, profileNode, "enableAutoJoin", vbNullString
        End If
        
        If profile.enableAutoIdentify Then
            xmlAddElement xml, profileNode, "enableAutoIdentify", vbNullString
        End If
        
        If profile.enableReconnect Then
            xmlAddElement xml, profileNode, "enableReconnect", vbNullString
        End If
        
        If profile.enableConnectRetry Then
            xmlAddElement xml, profileNode, "enableConnectRetry", vbNullString
        End If
        
        If profile.autoJoinChannelCount > 0 Then
            Dim channelsNode As IXMLDOMNode
            Dim channelNode As IXMLDOMNode
            
            Set channelsNode = xml.createNode(NODE_ELEMENT, "autoJoinChannels", "")
            profileNode.appendChild channelsNode
            
            Dim count As Long
            
            For count = 1 To profile.autoJoinChannelCount
                Set channelNode = xml.createNode(NODE_ELEMENT, "channel", "")
                channelsNode.appendChild channelNode
                
                xmlAddElement xml, channelNode, "name", profile.autoJoinChannel(count).channel
                
                If LenB(profile.autoJoinChannel(count).key) <> 0 Then
                    xmlAddElement xml, channelNode, "key", profile.autoJoinChannel(count).key
                End If
            Next count
        End If
    Next profile
    
    On Error GoTo ProfileSave_Error
    
    saveXml xml, filename
    RaiseEvent profilesChanged
Exit Sub
ProfileSave_Error:
    handleError "saveProfile", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub Class_Initialize()
    Set m_quickConnectProfile = New CServerProfile
    
    m_quickConnectProfile.name = "Quick connect"
    m_quickConnectProfile.hostname = "irc.swiftirc.net"
    m_quickConnectProfile.port = 6667
    m_quickConnectProfile.enableConnectRetry = True
    m_quickConnectProfile.enableReconnect = True
End Sub
