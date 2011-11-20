VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eSettingType
    estString
    estNumber
    estBoolean
End Enum

Private m_acceptedFirstUse As Boolean

Private m_quickConnectNickname As String
Private m_quickConnectChannel As String
Private m_quickConnectPassword As String

Private m_fontName As String
Private m_fontSize As Long

Private m_enableFiltering As Boolean
Private m_autoRejoinOnKick As Boolean

Private m_enableHighlighting As Boolean
Private m_highlightNickname As Boolean

Private m_settings As New Collection

Public Property Get setting(name As String, Optional settingType As eSettingType) As Variant
          Dim settingItem As CSettingItem
          
10        Set settingItem = findSetting(name)
          
20        If Not settingItem Is Nothing Then
30            If settingType = estString Then
40                setting = settingItem.value
50            ElseIf settingType = estNumber Then
60                setting = Val(settingItem.value)
70            ElseIf settingType = estBoolean Then
80                setting = textToBool(settingItem.value)
90            End If
100       Else
110           setting = False
120       End If
End Property

Public Property Let setting(name As String, Optional settingType As eSettingType, value As Variant)
          Dim settingItem As CSettingItem
          
10        Set settingItem = findSetting(name)
          
20        If settingItem Is Nothing Then
30            Set settingItem = New CSettingItem
40            settingItem.name = name
50            m_settings.Add settingItem, LCase$(name)
60        End If
          
70        If settingType = estString Then
80            settingItem.value = CStr(value)
90        ElseIf settingType = estNumber Then
100           settingItem.value = CStr(value)
110       ElseIf settingType = estBoolean Then
120           settingItem.value = boolToText(CBool(value))
130       End If
End Property

Private Function findSetting(name As String) As CSettingItem
10        On Error Resume Next
20        Set findSetting = m_settings.item(LCase$(name))
End Function

Public Property Get autoRejoinOnKick() As Boolean
10        autoRejoinOnKick = m_autoRejoinOnKick
End Property

Public Property Let autoRejoinOnKick(newValue As Boolean)
10        m_autoRejoinOnKick = newValue
End Property

Public Property Get enableFiltering() As Boolean
10        enableFiltering = m_enableFiltering
End Property

Public Property Let enableFiltering(newValue As Boolean)
10        m_enableFiltering = newValue
End Property

Public Property Get acceptedFirstUse() As Boolean
10        acceptedFirstUse = m_acceptedFirstUse
End Property

Public Property Let acceptedFirstUse(newValue As Boolean)
10        m_acceptedFirstUse = newValue
End Property

Public Property Get enableHighlighting() As Boolean
10        enableHighlighting = m_enableHighlighting
End Property

Public Property Let enableHighlighting(newValue As Boolean)
10        m_enableHighlighting = newValue
End Property

Public Property Get highlightNickname() As Boolean
10        highlightNickname = m_highlightNickname
End Property

Public Property Let highlightNickname(newValue As Boolean)
10        m_highlightNickname = newValue
End Property

Public Property Get fontName() As String
10        fontName = m_fontName
End Property

Public Property Let fontName(newValue As String)
10        m_fontName = newValue
End Property

Public Property Get fontSize() As Long
10        fontSize = m_fontSize
End Property

Public Property Let fontSize(newValue As Long)
10        m_fontSize = newValue
End Property

Public Property Get quickConnectNickname() As String
10        quickConnectNickname = m_quickConnectNickname
End Property

Public Property Let quickConnectNickname(newValue As String)
10        m_quickConnectNickname = newValue
End Property

Public Property Get quickConnectChannel() As String
10        quickConnectChannel = m_quickConnectChannel
End Property

Public Property Let quickConnectChannel(newValue As String)
10        m_quickConnectChannel = newValue
End Property

Public Property Get quickConnectPassword() As String
10        quickConnectPassword = m_quickConnectPassword
End Property

Public Property Let quickConnectPassword(newValue As String)
10        m_quickConnectPassword = newValue
End Property

Public Sub loadSettings()
          Dim xml As New DOMDocument30
          Dim node As IXMLDOMNode
          
10        If Not xml.load(g_userPath & "swiftirc_settings.xml") Then
20            Exit Sub
30        End If
          
40        m_quickConnectNickname = xmlGetElementText(xml.documentElement, "quickConnectNickname")
50        m_quickConnectChannel = xmlGetElementText(xml.documentElement, "quickConnectChannel")
60        m_quickConnectPassword = xmlGetElementText(xml.documentElement, "quickConnectPassword")
          
70        If LenB(m_quickConnectPassword) <> 0 Then
80            m_quickConnectPassword = decrypt(g_cryptKey, m_quickConnectPassword)
90        End If
          
100       m_fontName = xmlGetElementText(xml.documentElement, "fontName")
110       m_fontSize = Val(xmlGetElementText(xml.documentElement, "fontSize"))
          
120       m_enableHighlighting = textToBool(xmlGetElementText(xml.documentElement, "enableHighlighting"))
130       m_highlightNickname = textToBool(xmlGetElementText(xml.documentElement, "highlightNickname"))
          
140       m_acceptedFirstUse = xmlElementExists(xml.documentElement, "acceptedFirstUse")
          
150       If xmlElementExists(xml.documentElement, "disableFiltering") Then
160           m_enableFiltering = False
170       End If
          
180       m_autoRejoinOnKick = xmlElementExists(xml.documentElement, "autoRejoinOnKick")
          
          Dim settingItem As CSettingItem
          
190       For Each node In xml.documentElement.childNodes
200           Set settingItem = findSetting(node.nodeName)
              
210           If settingItem Is Nothing Then
220               Set settingItem = New CSettingItem
                  
230               settingItem.name = node.nodeName
240               settingItem.value = node.text
                  
250               m_settings.Add settingItem, LCase$(settingItem.name)
260           Else
270               settingItem.value = node.text
280           End If
290       Next node
          
300       g_timestamps = setting("timestamps", estBoolean)
310       g_timestampFormat = setting("timestampFormat", estString)
End Sub

Public Sub saveSettings()
          Dim xml As New DOMDocument30
          
10        Set xml.documentElement = xml.createNode(NODE_ELEMENT, "settings", "")
          
20        xmlAddElement xml, xml.documentElement, "quickConnectNickname", m_quickConnectNickname
30        xmlAddElement xml, xml.documentElement, "quickConnectChannel", m_quickConnectChannel
          
40        If LenB(m_quickConnectPassword) <> 0 Then
50            xmlAddElement xml, xml.documentElement, "quickConnectPassword", encrypt(g_cryptKey, _
                  m_quickConnectPassword)
60        End If
          
70        xmlAddElement xml, xml.documentElement, "fontName", m_fontName
80        xmlAddElement xml, xml.documentElement, "fontSize", CStr(m_fontSize)
          
90        xmlAddElement xml, xml.documentElement, "enableHighlighting", boolToText(m_enableHighlighting)
100       xmlAddElement xml, xml.documentElement, "highlightNickname", boolToText(m_highlightNickname)
          
110       If m_acceptedFirstUse Then
120           xmlAddElement xml, xml.documentElement, "acceptedFirstUse", vbNullString
130       End If
          
140       If Not m_enableFiltering Then
150           xmlAddElement xml, xml.documentElement, "disableFiltering", vbNullString
160       End If
          
170       If m_autoRejoinOnKick Then
180           xmlAddElement xml, xml.documentElement, "autoRejoinOnKick", vbNullString
190       End If
          
          Dim setting As CSettingItem
          
200       For Each setting In m_settings
210           xmlAddElement xml, xml.documentElement, setting.name, setting.value
220       Next setting
          
230       On Error GoTo SettingsSave_Error
          
240       saveXml xml, g_userPath & "swiftirc_settings.xml"
          
250   Exit Sub
SettingsSave_Error:
260       handleError "saveSettings", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub Class_Initialize()
10        m_enableFiltering = True
          
20        setting("switchbarColourEvent", estNumber) = 2
30        setting("switchbarColourMessage", estNumber) = 4
40        setting("switchbarColourAlert", estNumber) = 4
50        setting("switchbarColourHighlight", estNumber) = 3
60        setting("switchbarFlashAlert", estBoolean) = True
70        setting("switchbarFlashHighlight", estBoolean) = True
          
80        setting("nickColourOps", estNumber) = 4
90        setting("nickColourHalfOps", estNumber) = 7
100       setting("nickColourVoices", estNumber) = 12
110       setting("nickColourMe", estNumber) = 1
120       setting("nickColourNormal", estNumber) = 1
          
130       setting("nicknameIcons", estBoolean) = True
          
140       setting("timestamps", estBoolean) = True
150       setting("timestampFormat", estString) = "(hh:mm:ss)"
          
160       g_timestamps = True
170       g_timestampFormat = "(hh:mm:ss)"
End Sub
