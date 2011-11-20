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
10        Set urls = m_colourUrls
End Property

Public Property Get otherText() As CEventColour
10        Set otherText = m_colourOther
End Property

Public Property Get normalText() As CEventColour
10        Set normalText = m_colourNormal
End Property

Public Property Get noticeText() As CEventColour
10        Set noticeText = m_colourNotice
End Property

Public Property Get emotes() As CEventColour
10        Set emotes = m_colourEmote
End Property

Public Property Get channelJoin() As CEventColour
10        Set channelJoin = m_colourJoin
End Property

Public Property Get channelPart() As CEventColour
10        Set channelPart = m_colourPart
End Property

Public Property Get channelKick() As CEventColour
10        Set channelKick = m_colourKick
End Property

Public Property Get quit() As CEventColour
10        Set quit = m_colourQuit
End Property

Public Property Get modeChange() As CEventColour
10        Set modeChange = m_colourMode
End Property

Public Property Get topicChange() As CEventColour
10        Set topicChange = m_colourTopic
End Property

Public Property Get nickChanges() As CEventColour
10        Set nickChanges = m_colourNickChanges
End Property

Public Property Get whoisText() As CEventColour
10        Set whoisText = m_colourWhois
End Property

Public Property Get ctcpText() As CEventColour
10        Set ctcpText = m_colourCtcp
End Property

Public Property Get infoText() As CEventColour
10        Set infoText = m_colourInfo
End Property

Public Property Get myMessages() As CEventColour
10        Set myMessages = m_colourMyMessages
End Property

Public Property Get highlightText() As CEventColour
10        Set highlightText = m_colourHighlight
End Property

Private Sub initEventColours()
10        Set m_colourOther = addEventColour("Other text", "other", 1)
20        Set m_colourNormal = addEventColour("Normal messages", "normal", 1)
30        Set m_colourNotice = addEventColour("Notices", "notice", 5)
40        Set m_colourEmote = addEventColour("Emotes", "emote", 6)
50        Set m_colourJoin = addEventColour("Channel joins", "join", 3)
60        Set m_colourPart = addEventColour("Channel parts", "part", 3)
70        Set m_colourKick = addEventColour("Channel kicks", "kick", 3)
80        Set m_colourQuit = addEventColour("Quits", "quit", 2)
90        Set m_colourMode = addEventColour("Mode changes", "mode", 3)
100       Set m_colourTopic = addEventColour("Topic changes", "topic", 3)
110       Set m_colourNickChanges = addEventColour("Nickname changes", "nick", 3)
120       Set m_colourWhois = addEventColour("WHOIS text", "whois", 1)
130       Set m_colourCtcp = addEventColour("CTCP text", "ctcp", 4)
140       Set m_colourInfo = addEventColour("Client information", "info", 2)
          
150       Set m_colourMyMessages = addEventColour("My messages", "ownmessages", 1)
          
160       Set m_colourHighlight = addEventColour("Highlighted text", "highlight", 4)
170       Set m_colourUrls = addEventColour("Web addresses", "url", 2)
End Sub

Public Sub loadTheme(theme As CColourTheme)
10        m_colourOther.colour = theme.eventColour(1)
20        m_colourNormal.colour = theme.eventColour(2)
30        m_colourNotice.colour = theme.eventColour(3)
40        m_colourEmote.colour = theme.eventColour(4)
50        m_colourJoin.colour = theme.eventColour(5)
60        m_colourPart.colour = theme.eventColour(6)
70        m_colourKick.colour = theme.eventColour(7)
80        m_colourQuit.colour = theme.eventColour(8)
90        m_colourMode.colour = theme.eventColour(9)
100       m_colourTopic.colour = theme.eventColour(10)
110       m_colourNickChanges.colour = theme.eventColour(11)
120       m_colourWhois.colour = theme.eventColour(12)
130       m_colourCtcp.colour = theme.eventColour(13)
140       m_colourInfo.colour = theme.eventColour(14)
150       m_colourMyMessages.colour = theme.eventColour(15)
160       m_colourHighlight.colour = theme.eventColour(16)
170       m_colourUrls.colour = theme.eventColour(17)
End Sub

Public Function getName(index As Long) As String
10        getName = m_eventColours.item(index).name
End Function

Public Function getConfName(index As Long) As String
10        getConfName = m_eventColours.item(index).confName
End Function

Private Function addEventColour(name As String, confName As String, colour As Byte) As CEventColour
          Dim eventColour As New CEventColour
          
10        eventColour.name = name
20        eventColour.confName = confName
30        eventColour.colour = colour
40        m_eventColours.Add eventColour
          
50        Set addEventColour = eventColour
End Function

Public Sub copyEventColours(list As cArrayList)
          Dim count As Long
          Dim eventColour As CEventColour
          
10        For count = 1 To m_eventColours.count
20            Set eventColour = New CEventColour
              
30            eventColour.name = m_eventColours.item(count).name
40            eventColour.colour = m_eventColours.item(count).colour
50            list.Add eventColour
60        Next count
End Sub

Private Sub Class_Initialize()
10        initEventColours
End Sub
