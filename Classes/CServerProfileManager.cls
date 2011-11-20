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

Public Event profilesChanged()

Public Property Get quickConnectProfile() As CServerProfile
10        Set quickConnectProfile = m_quickConnectProfile
End Property

Public Sub addProfile(profile As CServerProfile)
10        If Not findProfile(profile.name) Is Nothing Then
20            Exit Sub
30        End If
          
40        m_serverProfiles.Add profile, LCase$(profile.name)
50        RaiseEvent profilesChanged
End Sub

Public Sub removeProfileIndex(index As Long)
10        m_serverProfiles.Remove index
20        RaiseEvent profilesChanged
End Sub

Public Sub removeProfile(name As String)
10        On Error Resume Next
20        m_serverProfiles.Remove LCase$(name)
30        RaiseEvent profilesChanged
End Sub

Public Function findProfile(name As String) As CServerProfile
10        On Error Resume Next
20        Set findProfile = m_serverProfiles.item(LCase$(name))
End Function

Public Function profileItem(index As Long) As CServerProfile
10        Set profileItem = m_serverProfiles.item(index)
End Function

Public Function profileCount() As Long
10        profileCount = m_serverProfiles.count
End Function

Public Sub loadProfiles(filename As String)
          Dim xml As New DOMDocument30
          Dim profileNode As IXMLDOMNode
          
10        If Not xml.load(filename) Then
20            Exit Sub
30        End If
          
          Dim profile As CServerProfile
          Dim channelsNode As IXMLDOMNode
          Dim channelNode As IXMLDOMNode
          
40        For Each profileNode In xml.documentElement.childNodes
50            Set profile = New CServerProfile
              
60            profile.name = profileNode.Attributes.getNamedItem("name").text
70            profile.hostname = xmlGetElementText(profileNode, "hostname")
80            profile.port = Val(xmlGetElementText(profileNode, "port"))
90            profile.serverPassword = xmlGetElementText(profileNode, "serverPassword")
100           profile.primaryNickname = xmlGetElementText(profileNode, "primaryNickname")
110           profile.backupNickname = xmlGetElementText(profileNode, "backupNickname")
120           profile.realName = xmlGetElementText(profileNode, "realName")
              
130           If xmlGetAttributeText(profileNode, "nicknamePassword", "crypt") = "yes" Then
140               profile.nicknamePassword = decrypt(g_cryptKey, xmlGetElementText(profileNode, "nicknamePassword"))
150           Else
160               profile.nicknamePassword = xmlGetElementText(profileNode, "nicknamePassword")
170           End If
              
180           profile.enablePerform = xmlElementExists(profileNode, "enablePerform")
190           profile.perform = Replace(xmlGetElementText(profileNode, "perform"), "\n", vbCrLf)
              
200           profile.enableAutoJoin = xmlElementExists(profileNode, "enableAutoJoin")
210           profile.enableAutoIdentify = xmlElementExists(profileNode, "enableAutoIdentify")
              
220           profile.enableReconnect = xmlElementExists(profileNode, "enableReconnect")
230           profile.enableConnectRetry = xmlElementExists(profileNode, "enableConnectRetry")
              
240           Set channelsNode = profileNode.selectSingleNode("autoJoinChannels")
              
250           If Not channelsNode Is Nothing Then
260               For Each channelNode In channelsNode.childNodes
270                   profile.addAutoJoinChannel xmlGetElementText(channelNode, "name"), _
                          xmlGetElementText(channelNode, "key")
280               Next channelNode
290           End If
              
300           addProfile profile
310       Next profileNode
          
320       RaiseEvent profilesChanged
End Sub

Public Sub saveProfiles(filename As String)
          Dim profile As CServerProfile
          Dim xml As New DOMDocument30
          Dim profileNode As IXMLDOMNode
          Dim node As IXMLDOMNode
          Dim attrib As IXMLDOMAttribute
          
10        Set xml.documentElement = xml.createNode(NODE_ELEMENT, "profiles", "")
          
20        For Each profile In m_serverProfiles
30            Set profileNode = xml.createNode(NODE_ELEMENT, "profile", "")
40            Set attrib = xml.createNode(NODE_ATTRIBUTE, "name", "")
              
50            attrib.text = profile.name
60            profileNode.Attributes.setNamedItem attrib
70            xml.documentElement.appendChild profileNode
              
80            xmlAddElement xml, profileNode, "hostname", profile.hostname
              
90            If profile.port <> 0 Then
100               xmlAddElement xml, profileNode, "port", profile.port
110           End If
              
120           If LenB(profile.serverPassword) <> 0 Then
130               xmlAddElement xml, profileNode, "serverPassword", profile.serverPassword
140           End If
              
150           xmlAddElement xml, profileNode, "primaryNickname", profile.primaryNickname
              
160           If LenB(profile.backupNickname) <> 0 Then
170               xmlAddElement xml, profileNode, "backupNickname", profile.backupNickname
180           End If
              
190           If LenB(profile.realName) <> 0 Then
200               xmlAddElement xml, profileNode, "realName", profile.realName
210           End If
              
220           If LenB(profile.nicknamePassword) <> 0 Then
                  Dim passNode As IXMLDOMNode
230               Set passNode = xmlAddElement(xml, profileNode, "nicknamePassword", encrypt(g_cryptKey, profile.nicknamePassword))
                  
240               Set attrib = xml.createAttribute("crypt")
250               attrib.text = "yes"
260               passNode.Attributes.setNamedItem attrib
270           End If
              
280           If profile.enablePerform Then
290               xmlAddElement xml, profileNode, "enablePerform", vbNullString
300           End If
              
310           If LenB(profile.perform) <> 0 Then
320               xmlAddElement xml, profileNode, "perform", Replace(profile.perform, vbCrLf, "\n")
330           End If
              
340           If profile.enableAutoJoin Then
350               xmlAddElement xml, profileNode, "enableAutoJoin", vbNullString
360           End If
              
370           If profile.enableAutoIdentify Then
380               xmlAddElement xml, profileNode, "enableAutoIdentify", vbNullString
390           End If
              
400           If profile.enableReconnect Then
410               xmlAddElement xml, profileNode, "enableReconnect", vbNullString
420           End If
              
430           If profile.enableConnectRetry Then
440               xmlAddElement xml, profileNode, "enableConnectRetry", vbNullString
450           End If
              
460           If profile.autoJoinChannelCount > 0 Then
                  Dim channelsNode As IXMLDOMNode
                  Dim channelNode As IXMLDOMNode
                  
470               Set channelsNode = xml.createNode(NODE_ELEMENT, "autoJoinChannels", "")
480               profileNode.appendChild channelsNode
                  
                  Dim count As Long
                  
490               For count = 1 To profile.autoJoinChannelCount
500                   Set channelNode = xml.createNode(NODE_ELEMENT, "channel", "")
510                   channelsNode.appendChild channelNode
                      
520                   xmlAddElement xml, channelNode, "name", profile.autoJoinChannel(count).channel
                      
530                   If LenB(profile.autoJoinChannel(count).key) <> 0 Then
540                       xmlAddElement xml, channelNode, "key", profile.autoJoinChannel(count).key
550                   End If
560               Next count
570           End If
580       Next profile
          
590       On Error GoTo ProfileSave_Error
          
600       saveXml xml, filename
610       RaiseEvent profilesChanged
620   Exit Sub
ProfileSave_Error:
630       handleError "saveProfile", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub Class_Initialize()
10        Set m_quickConnectProfile = New CServerProfile
          
20        m_quickConnectProfile.name = "Quick connect"
30        m_quickConnectProfile.hostname = "irc.swiftirc.net"
40        m_quickConnectProfile.port = 6667
50        m_quickConnectProfile.enableConnectRetry = True
60        m_quickConnectProfile.enableReconnect = True
End Sub
