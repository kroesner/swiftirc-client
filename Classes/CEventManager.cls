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
10        Set m_defaultTheme = newValue
End Property

Private Sub addEvent(name As String, format As String, eventColour As CEventColour, flags As Long)
          Dim newEvent As New CEvent
          
10        newEvent.init name, format, eventColour, flags
20        m_events.Add newEvent, LCase$(name)
End Sub

Private Sub defEvent(name As String)
10        addEvent name, vbNullString, Nothing, 0
End Sub

Public Sub loadTheme(theme As CTextTheme)
10        If Not theme Is m_defaultTheme Then
20            realLoadTheme m_defaultTheme
30        End If
          
40        realLoadTheme theme
50        RaiseEvent themeUpdated
End Sub

Private Sub realLoadTheme(theme As CTextTheme)
          Dim count  As Long
          Dim aEvent As CEvent
          
10        For count = 1 To theme.eventCount
20            Set aEvent = theme.getEvent(count)
30            updateEvent aEvent.eventName, aEvent.format, aEvent.eventColour, aEvent.flags
40        Next count
End Sub

Public Sub updateEvent(name As String, format As String, eventColour As CEventColour, flags As Long)
          Dim aEvent As CEvent
          
10        Set aEvent = findEvent(name)
          
20        If Not aEvent Is Nothing Then
30            aEvent.update format, eventColour, flags
40        End If
End Sub

Private Sub initEvents()
10        addEvent "CONNECTING", vbNullString, Nothing, 0
20        addEvent "CONNECTED", vbNullString, Nothing, 0
30        addEvent "DISCONNECTED", vbNullString, Nothing, 0
40        addEvent "RECONNECTING_IN", vbNullString, Nothing, 0

50        defEvent "IRC_ERROR"
60        defEvent "KILLED"

70        addEvent "NICKNAME_IN_USE", vbNullString, Nothing, 0
80        addEvent "NICKNAME_IN_USE_PREREG", vbNullString, Nothing, 0
90        addEvent "NICKNAME_IN_USE_PREREG2", vbNullString, Nothing, 0
100       defEvent "NO_SUCH_NICK"

110       addEvent "WELCOME", vbNullString, Nothing, 0
120       addEvent "NUMERIC", vbNullString, Nothing, 0

130       addEvent "PRIVMSG", vbNullString, Nothing, 0
140       addEvent "CHANNEL_PRIVMSG", vbNullString, Nothing, 0
150       addEvent "WALLCHOP_PRIVMSG", vbNullString, Nothing, 0
160       addEvent "ME_PRIVMSG", vbNullString, Nothing, 0
170       addEvent "NOTICE", vbNullString, Nothing, 0
180       addEvent "CHANNEL_NOTICE", vbNullString, Nothing, 0
190       addEvent "WALLCHOP_NOTICE", vbNullString, Nothing, 0
200       addEvent "EMOTE", vbNullString, Nothing, 0
210       addEvent "CHANNEL_EMOTE", vbNullString, Nothing, 0
220       defEvent "WALLCHOP_EMOTE"
          
230       addEvent "CTCP_RECEIVED", vbNullString, Nothing, 0
240       defEvent "CTCP_REPLY_RECEIVED"
          
250       addEvent "ME_JOIN", vbNullString, Nothing, 0
260       addEvent "ME_PART", vbNullString, Nothing, 0
270       addEvent "ME_REJOINING", vbNullString, Nothing, 0
280       addEvent "ME_REJOINED", vbNullString, Nothing, 0
290       defEvent "ME_REJOIN_DELAY"
          
300       defEvent "ME_MODE_CHANGE"
          
310       defEvent "ME_KICKED"
320       addEvent "USER_JOIN", vbNullString, Nothing, 0
330       addEvent "USER_PART", vbNullString, Nothing, 0
340       addEvent "USER_PART_REASON", vbNullString, Nothing, 0
350       defEvent "USER_QUIT"
360       addEvent "USER_QUIT_EX", vbNullString, Nothing, 0
370       addEvent "USER_QUIT_REASON", vbNullString, Nothing, 0
380       defEvent "USER_QUIT_REASON_EX"
390       addEvent "USER_NICK_CHANGE", vbNullString, Nothing, 0
400       defEvent "USER_KICKED"
          
410       defEvent "CHANNEL_PRIVMSG_HIGHLIGHT"
420       defEvent "CHANNEL_EMOTE_HIGHLIGHT"
430       defEvent "EMOTE_HIGHLIGHT"
440       defEvent "PRIVMSG_HIGHLIGHT"
          
450       defEvent "CHANNEL_MODE_CHANGE"
          
460       defEvent "CHANNEL_CTCP"
470       defEvent "WALLCHOP_CTCP"
480       defEvent "CTCP"
          
490       addEvent "CHANNEL_TOPICIS", vbNullString, Nothing, 0
500       addEvent "CHANNEL_TOPICWHOTIME", vbNullString, Nothing, 0
510       defEvent "CHANNEL_TOPICCHANGE"
          
520       addEvent "ERROR_CONNECT", vbNullString, Nothing, 0
530       addEvent "ERROR_DISCONNECT", vbNullString, Nothing, 0
          
540       defEvent "WHOIS_USER"
550       defEvent "WHOIS_CHANNELS"
560       defEvent "WHOIS_SERVER"
570       defEvent "WHOIS_OPERATOR"
580       defEvent "WHOIS_IDLE"
590       defEvent "WHOIS_REGNICK"
600       defEvent "WHOIS_GENERIC"
610       defEvent "WHOIS_END"
620       defEvent "AWAY"
          
630       defEvent "ERR_BANNEDFROMCHAN"
640       defEvent "ERR_INVITEONLYCHAN"
650       defEvent "ERR_NEEDREGGEDNICK"
660       defEvent "ERR_NOPRIVILEGES"
670       defEvent "ERR_BADCHANNELKEY"
680       defEvent "ERR_CHANNELISFULL"
690       defEvent "ERR_TOOMANYJOINS"
700       defEvent "ERR_TOOMANYCHANNELS"
710       defEvent "ERR_SECUREONLYCHAN"
          
720       defEvent "CMD_INSUFFICIENT_PARAMS"
730       defEvent "CMD_INCOMPATIBLE_WINDOW"
          
740       defEvent "CMD_PRIVMSG_SENT"
750       defEvent "CMD_NOTICE_SENT"
760       defEvent "CMD_CTCP_SENT"
770       defEvent "CMD_RAW_SENT"
          
780       defEvent "WHO_LIST"
790       defEvent "WHO_END_OF_LIST"
800       defEvent "WHOWAS_HOST"
810       defEvent "WHOWAS_UNKNOWN"
820       defEvent "WHOWAS_END"
830       defEvent "SILENCE_LIST"
840       defEvent "END_OF_SILENCE_LIST"
850       defEvent "ERR_CHGNICK_MODEN"
860       defEvent "ERR_CHGNICK_MODEB"
870       defEvent "HELPOP_TITLE"
880       defEvent "HELPOP_TEXT"
890       defEvent "MARKED_AWAY"
          
900       defEvent "IGNORE_HELP"
910       defEvent "IGNORE_HELP_FLAGS"
          
920       defEvent "IGNORE_LIST_START"
930       defEvent "IGNORE_LIST_ENTRY"
940       defEvent "IGNORE_LIST_END"
          
950       defEvent "IGNORE_ADDED"
960       defEvent "IGNORE_UPDATED"
970       defEvent "IGNORE_REMOVED"
980       defEvent "IGNORE_REMOVE_NOTFOUND"
990       defEvent "IGNORE_LIST_CLEARED"
          
1000      defEvent "IGNORE_INVALID_FLAGS"
1010      defEvent "IGNORE_INVALID_COMMAND"
1020      defEvent "NO_LONGER_MARKED_AWAY"
1030      defEvent "INVITE_LIST"
1040      defEvent "END_OF_INVITE_LIST"
1050      defEvent "INVITE_USER"
          
1060      defEvent "BAN_LIST"
1070      defEvent "END_OF_BAN_LIST"
1080      defEvent "EX_LIST"
1090      defEvent "END_OF_EX_LIST"
1100      defEvent "INVEX_LIST"
1110      defEvent "END_OF_INVEX_LIST"
1120      defEvent "A_LIST"
1130      defEvent "END_OF_A_LIST"
1140      defEvent "Q_LIST"
1150      defEvent "END_OF_Q_LIST"
          
1160      defEvent "SILENCE_MODIFY"
1170      defEvent "INVITATION_RECEIVED"
End Sub

Private Sub loadThemes()

End Sub

Public Function findEvent(name As String) As CEvent
10        On Error Resume Next
20        Set findEvent = m_events.item(LCase$(name))
End Function

Private Sub Class_Initialize()
10        initEvents
End Sub
