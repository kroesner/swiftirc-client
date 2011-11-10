VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_events As New Collection
Private m_defaultTheme As CTextTheme

Public Event themeUpdated()

Public Property Let defaultTheme(newValue As CTextTheme)
    Set m_defaultTheme = newValue
End Property

Private Sub addEvent(name As String, format As String, eventColour As CEventColour, flags As Long)
    Dim newEvent As New CEvent
    
    newEvent.init name, format, eventColour, flags
    m_events.Add newEvent, LCase$(name)
End Sub

Private Sub defEvent(name As String)
    addEvent name, vbNullString, Nothing, 0
End Sub

Public Sub loadTheme(theme As CTextTheme)
    If Not theme Is m_defaultTheme Then
        realLoadTheme m_defaultTheme
    End If
    
    realLoadTheme theme
    RaiseEvent themeUpdated
End Sub

Private Sub realLoadTheme(theme As CTextTheme)
    Dim count  As Long
    Dim aEvent As CEvent
    
    For count = 1 To theme.eventCount
        Set aEvent = theme.getEvent(count)
        updateEvent aEvent.eventName, aEvent.format, aEvent.eventColour, aEvent.flags
    Next count
End Sub

Public Sub updateEvent(name As String, format As String, eventColour As CEventColour, flags As Long)
    Dim aEvent As CEvent
    
    Set aEvent = findEvent(name)
    
    If Not aEvent Is Nothing Then
        aEvent.update format, eventColour, flags
    End If
End Sub

Private Sub initEvents()
    addEvent "CONNECTING", vbNullString, Nothing, 0
    addEvent "CONNECTED", vbNullString, Nothing, 0
    addEvent "DISCONNECTED", vbNullString, Nothing, 0
    addEvent "RECONNECTING_IN", vbNullString, Nothing, 0

    defEvent "IRC_ERROR"
    defEvent "KILLED"

    addEvent "NICKNAME_IN_USE", vbNullString, Nothing, 0
    addEvent "NICKNAME_IN_USE_PREREG", vbNullString, Nothing, 0
    addEvent "NICKNAME_IN_USE_PREREG2", vbNullString, Nothing, 0
    defEvent "NO_SUCH_NICK"

    addEvent "WELCOME", vbNullString, Nothing, 0
    addEvent "NUMERIC", vbNullString, Nothing, 0

    addEvent "PRIVMSG", vbNullString, Nothing, 0
    addEvent "CHANNEL_PRIVMSG", vbNullString, Nothing, 0
    addEvent "WALLCHOP_PRIVMSG", vbNullString, Nothing, 0
    addEvent "ME_PRIVMSG", vbNullString, Nothing, 0
    addEvent "NOTICE", vbNullString, Nothing, 0
    addEvent "CHANNEL_NOTICE", vbNullString, Nothing, 0
    addEvent "WALLCHOP_NOTICE", vbNullString, Nothing, 0
    addEvent "EMOTE", vbNullString, Nothing, 0
    addEvent "CHANNEL_EMOTE", vbNullString, Nothing, 0
    defEvent "WALLCHOP_EMOTE"
    
    addEvent "CTCP_RECEIVED", vbNullString, Nothing, 0
    defEvent "CTCP_REPLY_RECEIVED"
    
    addEvent "ME_JOIN", vbNullString, Nothing, 0
    addEvent "ME_PART", vbNullString, Nothing, 0
    addEvent "ME_REJOINING", vbNullString, Nothing, 0
    addEvent "ME_REJOINED", vbNullString, Nothing, 0
    defEvent "ME_REJOIN_DELAY"
    
    defEvent "ME_MODE_CHANGE"
    
    defEvent "ME_KICKED"
    addEvent "USER_JOIN", vbNullString, Nothing, 0
    addEvent "USER_PART", vbNullString, Nothing, 0
    addEvent "USER_PART_REASON", vbNullString, Nothing, 0
    defEvent "USER_QUIT"
    addEvent "USER_QUIT_EX", vbNullString, Nothing, 0
    addEvent "USER_QUIT_REASON", vbNullString, Nothing, 0
    defEvent "USER_QUIT_REASON_EX"
    addEvent "USER_NICK_CHANGE", vbNullString, Nothing, 0
    defEvent "USER_KICKED"
    
    defEvent "CHANNEL_PRIVMSG_HIGHLIGHT"
    defEvent "CHANNEL_EMOTE_HIGHLIGHT"
    defEvent "EMOTE_HIGHLIGHT"
    defEvent "PRIVMSG_HIGHLIGHT"
    
    defEvent "CHANNEL_MODE_CHANGE"
    
    defEvent "CHANNEL_CTCP"
    defEvent "WALLCHOP_CTCP"
    defEvent "CTCP"
    
    addEvent "CHANNEL_TOPICIS", vbNullString, Nothing, 0
    addEvent "CHANNEL_TOPICWHOTIME", vbNullString, Nothing, 0
    defEvent "CHANNEL_TOPICCHANGE"
    
    addEvent "ERROR_CONNECT", vbNullString, Nothing, 0
    addEvent "ERROR_DISCONNECT", vbNullString, Nothing, 0
    
    defEvent "WHOIS_USER"
    defEvent "WHOIS_CHANNELS"
    defEvent "WHOIS_SERVER"
    defEvent "WHOIS_OPERATOR"
    defEvent "WHOIS_IDLE"
    defEvent "WHOIS_REGNICK"
    defEvent "WHOIS_GENERIC"
    defEvent "WHOIS_END"
    defEvent "AWAY"
    
    defEvent "ERR_BANNEDFROMCHAN"
    defEvent "ERR_INVITEONLYCHAN"
    defEvent "ERR_NEEDREGGEDNICK"
    defEvent "ERR_NOPRIVILEGES"
    defEvent "ERR_BADCHANNELKEY"
    defEvent "ERR_CHANNELISFULL"
    defEvent "ERR_TOOMANYJOINS"
    defEvent "ERR_TOOMANYCHANNELS"
    defEvent "ERR_SECUREONLYCHAN"
    
    defEvent "CMD_INSUFFICIENT_PARAMS"
    defEvent "CMD_INCOMPATIBLE_WINDOW"
    
    defEvent "CMD_PRIVMSG_SENT"
    defEvent "CMD_NOTICE_SENT"
    defEvent "CMD_CTCP_SENT"
    
    ' 10/nov/2011
    defEvent "WHO_LIST"
    defEvent "WHO_END_OF_LIST"
    
    defEvent "WHOWAS_HOST"
    defEvent "WHOWAS_UNKNOWN"
    defEvent "WHOWAS_END"
    
    defEvent "SILENCE_LIST"
    defEvent "END_OF_SILENCE_LIST"
    defEvent "ERR_CHGNICK_MODEN"
    defEvent "ERR_CHGNICK_MODEB"
    defEvent "HELPOP_TITLE"
    defEvent "HELPOP_TEXT"
    defEvent "MARKED_AWAY"
    defEvent "NO_LONGER_MARKED_AWAY"
    defEvent "INVITE_LIST"
    defEvent "END_OF_INVITE_LIST"
    defEvent "INVITE_USER"
End Sub

Private Sub loadThemes()

End Sub

Public Function findEvent(name As String) As CEvent
    On Error Resume Next
    Set findEvent = m_events.item(LCase$(name))
End Function

Private Sub Class_Initialize()
    initEvents
End Sub
