VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventColours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_eventColours As New cArrayList

Private m_colourOther As CEventColour
Private m_colourNormal As CEventColour
Private m_colourNotice As CEventColour
Private m_colourEmote As CEventColour

Private m_colourJoin As CEventColour
Private m_colourPart As CEventColour
Private m_colourKick As CEventColour
Private m_colourQuit As CEventColour

Private m_colourMode As CEventColour
Private m_colourTopic As CEventColour
Private m_colourNickChanges As CEventColour

Private m_colourWhois As CEventColour
Private m_colourCtcp As CEventColour

Private m_colourInfo As CEventColour

Private m_colourMyMessages As CEventColour

Private m_colourHighlight As CEventColour

Private m_colourUrls As CEventColour

Public Property Get urls() As CEventColour
    Set urls = m_colourUrls
End Property

Public Property Get otherText() As CEventColour
    Set otherText = m_colourOther
End Property

Public Property Get normalText() As CEventColour
    Set normalText = m_colourNormal
End Property

Public Property Get noticeText() As CEventColour
    Set noticeText = m_colourNotice
End Property

Public Property Get emotes() As CEventColour
    Set emotes = m_colourEmote
End Property

Public Property Get channelJoin() As CEventColour
    Set channelJoin = m_colourJoin
End Property

Public Property Get channelPart() As CEventColour
    Set channelPart = m_colourPart
End Property

Public Property Get channelKick() As CEventColour
    Set channelKick = m_colourKick
End Property

Public Property Get quit() As CEventColour
    Set quit = m_colourQuit
End Property

Public Property Get modeChange() As CEventColour
    Set modeChange = m_colourMode
End Property

Public Property Get topicChange() As CEventColour
    Set topicChange = m_colourTopic
End Property

Public Property Get nickChanges() As CEventColour
    Set nickChanges = m_colourNickChanges
End Property

Public Property Get whoisText() As CEventColour
    Set whoisText = m_colourWhois
End Property

Public Property Get ctcpText() As CEventColour
    Set ctcpText = m_colourCtcp
End Property

Public Property Get infoText() As CEventColour
    Set infoText = m_colourInfo
End Property

Public Property Get myMessages() As CEventColour
    Set myMessages = m_colourMyMessages
End Property

Public Property Get highlightText() As CEventColour
    Set highlightText = m_colourHighlight
End Property

Private Sub initEventColours()
    Set m_colourOther = addEventColour("Other text", "other", 1)
    Set m_colourNormal = addEventColour("Normal messages", "normal", 1)
    Set m_colourNotice = addEventColour("Notices", "notice", 5)
    Set m_colourEmote = addEventColour("Emotes", "emote", 6)
    Set m_colourJoin = addEventColour("Channel joins", "join", 3)
    Set m_colourPart = addEventColour("Channel parts", "part", 3)
    Set m_colourKick = addEventColour("Channel kicks", "kick", 3)
    Set m_colourQuit = addEventColour("Quits", "quit", 2)
    Set m_colourMode = addEventColour("Mode changes", "mode", 3)
    Set m_colourTopic = addEventColour("Topic changes", "topic", 3)
    Set m_colourNickChanges = addEventColour("Nickname changes", "nick", 3)
    Set m_colourWhois = addEventColour("WHOIS text", "whois", 1)
    Set m_colourCtcp = addEventColour("CTCP text", "ctcp", 4)
    Set m_colourInfo = addEventColour("Client information", "info", 2)
    
    Set m_colourMyMessages = addEventColour("My messages", "ownmessages", 1)
    
    Set m_colourHighlight = addEventColour("Highlighted text", "highlight", 4)
    Set m_colourUrls = addEventColour("Web addresses", "url", 2)
End Sub

Public Sub loadTheme(theme As CColourTheme)
    m_colourOther.colour = theme.eventColour(1)
    m_colourNormal.colour = theme.eventColour(2)
    m_colourNotice.colour = theme.eventColour(3)
    m_colourEmote.colour = theme.eventColour(4)
    m_colourJoin.colour = theme.eventColour(5)
    m_colourPart.colour = theme.eventColour(6)
    m_colourKick.colour = theme.eventColour(7)
    m_colourQuit.colour = theme.eventColour(8)
    m_colourMode.colour = theme.eventColour(9)
    m_colourTopic.colour = theme.eventColour(10)
    m_colourNickChanges.colour = theme.eventColour(11)
    m_colourWhois.colour = theme.eventColour(12)
    m_colourCtcp.colour = theme.eventColour(13)
    m_colourInfo.colour = theme.eventColour(14)
    m_colourMyMessages.colour = theme.eventColour(15)
    m_colourHighlight.colour = theme.eventColour(16)
    m_colourUrls.colour = theme.eventColour(17)
End Sub

Public Function getName(index As Long) As String
    getName = m_eventColours.item(index).name
End Function

Public Function getConfName(index As Long) As String
    getConfName = m_eventColours.item(index).confName
End Function

Private Function addEventColour(name As String, confName As String, colour As Byte) As CEventColour
    Dim eventColour As New CEventColour
    
    eventColour.name = name
    eventColour.confName = confName
    eventColour.colour = colour
    m_eventColours.Add eventColour
    
    Set addEventColour = eventColour
End Function

Public Sub copyEventColours(list As cArrayList)
    Dim count As Long
    Dim eventColour As CEventColour
    
    For count = 1 To m_eventColours.count
        Set eventColour = New CEventColour
        
        eventColour.name = m_eventColours.item(count).name
        eventColour.colour = m_eventColours.item(count).colour
        list.Add eventColour
    Next count
End Sub

Private Sub Class_Initialize()
    initEventColours
End Sub
