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

Private Const SETTINGS_FILENAME  As String = "swiftirc_settings.xml"

Public Property Get setting(name As String, Optional settingType As eSettingType) As Variant
    Dim settingItem As CSettingItem
    
    Set settingItem = findSetting(name)
    
    If Not settingItem Is Nothing Then
        If settingType = estString Then
            setting = settingItem.value
        ElseIf settingType = estNumber Then
            setting = Val(settingItem.value)
        ElseIf settingType = estBoolean Then
            setting = textToBool(settingItem.value)
        End If
    Else
        setting = False
    End If
End Property

Public Property Let setting(name As String, Optional settingType As eSettingType, value As Variant)
    Dim settingItem As CSettingItem
    
    Set settingItem = findSetting(name)
    
    If settingItem Is Nothing Then
        Set settingItem = New CSettingItem
        settingItem.name = name
        m_settings.Add settingItem, LCase$(name)
    End If
    
    If settingType = estString Then
        settingItem.value = CStr(value)
    ElseIf settingType = estNumber Then
        settingItem.value = CStr(value)
    ElseIf settingType = estBoolean Then
        settingItem.value = boolToText(CBool(value))
    End If
End Property

Private Function findSetting(name As String) As CSettingItem
    On Error Resume Next
    Set findSetting = m_settings.item(LCase$(name))
End Function

Public Property Get autoRejoinOnKick() As Boolean
    autoRejoinOnKick = m_autoRejoinOnKick
End Property

Public Property Let autoRejoinOnKick(newValue As Boolean)
    m_autoRejoinOnKick = newValue
End Property

Public Property Get enableFiltering() As Boolean
    enableFiltering = m_enableFiltering
End Property

Public Property Let enableFiltering(newValue As Boolean)
    m_enableFiltering = newValue
End Property

Public Property Get acceptedFirstUse() As Boolean
    acceptedFirstUse = m_acceptedFirstUse
End Property

Public Property Let acceptedFirstUse(newValue As Boolean)
    m_acceptedFirstUse = newValue
End Property

Public Property Get enableHighlighting() As Boolean
    enableHighlighting = m_enableHighlighting
End Property

Public Property Let enableHighlighting(newValue As Boolean)
    m_enableHighlighting = newValue
End Property

Public Property Get highlightNickname() As Boolean
    highlightNickname = m_highlightNickname
End Property

Public Property Let highlightNickname(newValue As Boolean)
    m_highlightNickname = newValue
End Property

Public Property Get fontName() As String
    fontName = m_fontName
End Property

Public Property Let fontName(newValue As String)
    m_fontName = newValue
End Property

Public Property Get fontSize() As Long
    fontSize = m_fontSize
End Property

Public Property Let fontSize(newValue As Long)
    m_fontSize = newValue
End Property

Public Property Get quickConnectNickname() As String
    quickConnectNickname = m_quickConnectNickname
End Property

Public Property Let quickConnectNickname(newValue As String)
    m_quickConnectNickname = newValue
End Property

Public Property Get quickConnectChannel() As String
    quickConnectChannel = m_quickConnectChannel
End Property

Public Property Let quickConnectChannel(newValue As String)
    m_quickConnectChannel = newValue
End Property

Public Property Get quickConnectPassword() As String
    quickConnectPassword = m_quickConnectPassword
End Property

Public Property Let quickConnectPassword(newValue As String)
    m_quickConnectPassword = newValue
End Property

Public Sub loadSettings()
    Dim xml As New DOMDocument30
    Dim node As IXMLDOMNode
    
    If Not xml.load(combinePath(g_userPath, SETTINGS_FILENAME)) Then
        Exit Sub
    End If
    
    m_quickConnectNickname = xmlGetElementText(xml.documentElement, "quickConnectNickname")
    m_quickConnectChannel = xmlGetElementText(xml.documentElement, "quickConnectChannel")
    m_quickConnectPassword = xmlGetElementText(xml.documentElement, "quickConnectPassword")
    
    If LenB(m_quickConnectPassword) <> 0 Then
        m_quickConnectPassword = decrypt(g_cryptKey, m_quickConnectPassword)
    End If
    
    m_fontName = xmlGetElementText(xml.documentElement, "fontName")
    m_fontSize = Val(xmlGetElementText(xml.documentElement, "fontSize"))
    
    m_enableHighlighting = textToBool(xmlGetElementText(xml.documentElement, "enableHighlighting"))
    m_highlightNickname = textToBool(xmlGetElementText(xml.documentElement, "highlightNickname"))
    
    m_acceptedFirstUse = xmlElementExists(xml.documentElement, "acceptedFirstUse")
    
    If xmlElementExists(xml.documentElement, "disableFiltering") Then
        m_enableFiltering = False
    End If
    
    m_autoRejoinOnKick = xmlElementExists(xml.documentElement, "autoRejoinOnKick")
    
    Dim settingItem As CSettingItem
    
    For Each node In xml.documentElement.childNodes
        Set settingItem = findSetting(node.nodeName)
        
        If settingItem Is Nothing Then
            Set settingItem = New CSettingItem
            
            settingItem.name = node.nodeName
            settingItem.value = node.text
            
            m_settings.Add settingItem, LCase$(settingItem.name)
        Else
            settingItem.value = node.text
        End If
    Next node
    
    g_timestamps = setting("timestamps", estBoolean)
    g_timestampFormat = setting("timestampFormat", estString)
End Sub

Public Sub saveSettings()
    Dim xml As New DOMDocument30
    
    Set xml.documentElement = xml.createNode(NODE_ELEMENT, "settings", "")
    
    xmlAddElement xml, xml.documentElement, "quickConnectNickname", m_quickConnectNickname
    xmlAddElement xml, xml.documentElement, "quickConnectChannel", m_quickConnectChannel
    
    If LenB(m_quickConnectPassword) <> 0 Then
        xmlAddElement xml, xml.documentElement, "quickConnectPassword", encrypt(g_cryptKey, _
            m_quickConnectPassword)
    End If
    
    xmlAddElement xml, xml.documentElement, "fontName", m_fontName
    xmlAddElement xml, xml.documentElement, "fontSize", CStr(m_fontSize)
    
    xmlAddElement xml, xml.documentElement, "enableHighlighting", boolToText(m_enableHighlighting)
    xmlAddElement xml, xml.documentElement, "highlightNickname", boolToText(m_highlightNickname)
    
    If m_acceptedFirstUse Then
        xmlAddElement xml, xml.documentElement, "acceptedFirstUse", vbNullString
    End If
    
    If Not m_enableFiltering Then
        xmlAddElement xml, xml.documentElement, "disableFiltering", vbNullString
    End If
    
    If m_autoRejoinOnKick Then
        xmlAddElement xml, xml.documentElement, "autoRejoinOnKick", vbNullString
    End If
    
    Dim setting As CSettingItem
    
    For Each setting In m_settings
        xmlAddElement xml, xml.documentElement, setting.name, setting.value
    Next setting
    
    On Error GoTo SettingsSave_Error
    
    saveXml xml, combinePath(g_userPath, SETTINGS_FILENAME)
    
Exit Sub
SettingsSave_Error:
    handleError "saveSettings", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub Class_Initialize()
    m_enableFiltering = True
    
    setting("switchbarColourEvent", estNumber) = 2
    setting("switchbarColourMessage", estNumber) = 4
    setting("switchbarColourAlert", estNumber) = 4
    setting("switchbarColourHighlight", estNumber) = 3
    setting("switchbarFlashAlert", estBoolean) = True
    setting("switchbarFlashHighlight", estBoolean) = True
    
    setting("nickColourOps", estNumber) = 4
    setting("nickColourHalfOps", estNumber) = 7
    setting("nickColourVoices", estNumber) = 12
    setting("nickColourMe", estNumber) = 1
    setting("nickColourNormal", estNumber) = 1
    
    setting("nicknameIcons", estBoolean) = True
    
    setting("timestamps", estBoolean) = True
    setting("timestampFormat", estString) = "(hh:mm:ss)"
    
    g_timestamps = True
    g_timestampFormat = "(hh:mm:ss)"
End Sub
