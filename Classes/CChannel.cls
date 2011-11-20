VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_window As ctlWindowChannel
Private m_windowText As ITextWindow
Private m_session As CSession

Private m_isin As Boolean
Private m_hopping As Boolean
Private m_membersListSynced As Boolean
Private m_members As New Collection
Private m_listModes As New cArrayList
Private m_listSync As New Collection
Private m_paramModes As New cArrayList
Private m_modes As String

Private m_key As String
Private m_limit As Long
Private m_topic As String

Private m_topicHistory As New Collection
Private WithEvents m_rejoinTimer As VB.timer
Attribute m_rejoinTimer.VB_VarHelpID = -1

Public Event modeListSynced(mode As String)

Public Property Get meIsIn() As Boolean
10        meIsIn = m_isin
End Property

Public Property Get session() As CSession
10        Set session = m_session
End Property

Public Property Get topic() As String
10        topic = m_topic
End Property

Public Property Get limit() As Long
10        limit = m_limit
End Property

Public Property Get key() As String
10        key = m_key
End Property

Public Sub init(name As String, window As ctlWindowChannel, session As CSession)
10        m_name = name
20        Set m_window = window
30        Set m_windowText = window
40        Set m_session = session
          
50        Set m_rejoinTimer = m_session.client.getTimer
          
60        If settings.setting("enableLogging", estBoolean) And settings.setting("logChannel", estBoolean) Then
70            m_windowText.textview.logName = m_session.baseLogPath & sanitizeFilename(m_name)
80            m_windowText.textview.enableLogging = True
90        End If
End Sub

Public Sub deInit()
10        m_window.deInit

20        m_session.client.releaseTimer m_rejoinTimer

30        Set m_window = Nothing
40        Set m_windowText = Nothing
50        Set m_session = Nothing
End Sub

Public Property Get IText() As ITextWindow
10        Set IText = m_window
End Property

Public Property Get window() As ctlWindowChannel
10        Set window = m_window
End Property

Public Property Get name() As String
10        name = m_name
End Property

Public Property Get hopping() As Boolean
10        hopping = m_hopping
End Property

Public Property Let hopping(newValue As Boolean)
10        m_hopping = newValue
End Property

Public Sub disconnected()
10        clearMembers
20        m_windowText.addEvent "DISCONNECTED", makeStringArray("")
30        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
40        m_isin = False
End Sub

Public Sub reJoining()
10        m_windowText.addEvent "ME_REJOINING", makeStringArray(m_name)
20        clearMembers
30        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
40        m_isin = False
End Sub

Public Sub reJoined()
10        m_windowText.addEvent "ME_REJOINED", makeStringArray(m_name)
20        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
30        m_isin = True
End Sub

Public Sub meJoined()
10        m_windowText.addEvent "ME_JOIN", makeStringArray(m_name)
20        m_session.sendLine "MODE " & m_name
30        m_isin = True
          
40        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub hop(Optional reason As String = "Hopping")
10        m_hopping = True
20        m_session.sendLine "PART " & m_name & " :" & reason
End Sub

Public Sub userJoined(origin As COrigin)
10        If Not findMember(origin.nickname) Is Nothing Then
20            Exit Sub
30        End If
          
40        addChannelMember origin.nickname
50        m_windowText.addEvent "USER_JOIN", makeStringArray(origin.nickname, m_name, origin.username, _
              origin.hostname)
              
60        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userParted(origin As COrigin, Optional reason As String = vbNullString)
10        removeChannelMember origin.nickname
          
20        If LenB(reason) = 0 Then
30            m_windowText.addEvent "USER_PART", makeStringArray(origin.nickname, m_name, origin.username, _
                  origin.hostname)
40        Else
50            m_windowText.addEvent "USER_PART_REASON", makeStringArray(origin.nickname, m_name, reason, _
                  origin.username, origin.hostname)
60        End If
          
70        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userQuit(origin As COrigin, Optional reason As String = vbNullString)
10        If LenB(reason) = 0 Then
20            m_windowText.addEventEx "USER_QUIT", getUserStyle(origin.nickname), origin.nickname, _
                  TVE_NONE, makeStringArray("")
30        Else
40            m_windowText.addEventEx "USER_QUIT_REASON", getUserStyle(origin.nickname), origin.nickname, _
                  TVE_NONE, makeStringArray(reason)
50        End If
          
60        removeChannelMember origin.nickname
          
70        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userNickChange(origin As COrigin, newNick As String)
          Dim member As CChannelMember
          
10        Set member = findMember(origin.nickname)
          
20        If member Is Nothing Then
30            Exit Sub
40        End If
          
50        m_window.removeNicklistItem member.nickname, m_session.getEffectivePrefix(member.modes)
60        m_members.Remove LCase$(member.nickname)
70        member.nickname = newNick
80        m_members.Add member, LCase$(newNick)
90        m_window.addNicklistItem member.nickname, m_session.getEffectivePrefix(member.modes)
          
100       m_windowText.addEvent "USER_NICK_CHANGE", makeStringArray(origin.nickname, newNick)
          
110       m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userKicked(origin As COrigin, target As String, reason As String)
10        removeChannelMember target
20        m_windowText.addEvent "USER_KICKED", makeStringArray(origin.nickname, m_name, target, reason)
          
30        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub meKicked(origin As COrigin, reason As String)
10        clearMembers
20        m_isin = False
30        m_windowText.addEvent "ME_KICKED", makeStringArray(m_name, origin.nickname, reason)
          
40        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
          
50        If settings.autoRejoinOnKick Then
60            m_windowText.addEvent "ME_REJOIN_DELAY", makeStringArray(m_name, "2 seconds")
70            m_rejoinTimer.Interval = 2000
80            m_rejoinTimer.enabled = True
90        End If
End Sub

Private Sub addTopicHistory(topic As String)
          Dim tempTopic As Variant
          
10        For Each tempTopic In m_topicHistory
20            If StrComp(tempTopic, topic, vbBinaryCompare) = 0 Then
30                Exit Sub
40            End If
50        Next tempTopic
          
60        m_topicHistory.Add topic
End Sub

Public Function getTopicHistory(index As Long) As String
10        getTopicHistory = m_topicHistory.item(index)
End Function

Public Function getTopicHistoryCount() As Long
10        getTopicHistoryCount = m_topicHistory.count
End Function

Public Sub topicIs(topic As String)
10        m_windowText.addEvent "CHANNEL_TOPICIS", makeStringArray(topic)
20        m_topic = topic
          
30        addTopicHistory topic
End Sub

Public Sub topicWhoTime(who As String, time As Long)
10        m_windowText.addEvent "CHANNEL_TOPICWHOTIME", makeStringArray(who, formatTime(time))
End Sub

Public Sub topicChange(origin As COrigin, topic As String)
10        m_windowText.addEvent "CHANNEL_TOPICCHANGE", makeStringArray(origin.nickname, topic)
20        m_topic = topic
          
30        addTopicHistory topic
          
40        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub modeIs(modes As String, Optional paramString As String = vbNullString)
10        m_paramModes.clear
20        m_modes = vbNullString

          Dim count As Long
          Dim params() As String
          Dim mode As String
          
          Dim paramIndex As Long
          Dim paramCount As Long
          
          Dim modeType As eChannelModeType
                  
30        params = Split(paramString, " ")
40        paramCount = UBound(params) + 1
          
50        For count = 1 To Len(modes)
60            mode = Mid$(modes, count, 1)
70            modeType = m_session.getChannelModeType(mode)
              
80            Select Case modeType
                  Case cmtParam
90                    If paramIndex < paramCount Then
100                       addParamMode mode, params(paramIndex)
110                       paramIndex = paramIndex + 1
120                   End If
130               Case cmtSetOnly
140                   If paramIndex < paramCount Then
150                       addParamMode mode, params(paramIndex)
160                       paramIndex = paramIndex + 1
170                   End If
180               Case cmtNormal
190                   addMode mode
200           End Select
210       Next count
End Sub
    
Public Sub getModeList(mode As String, list As cArrayList)
          Dim count As Long
          
10        For count = 1 To m_listModes.count
20            If m_listModes.item(count).mode = mode Then
30                list.Add m_listModes.item(count)
40            End If
50        Next count
End Sub
    
Private Sub setListSync(mode As String, op As Boolean)
10        On Error Resume Next
          
20        If op Then
30            m_listSync.Add CVar(True), mode
40        Else
50            m_listSync.Remove mode
60        End If
End Sub

Public Function listIsSynced(mode As String) As Boolean
          'This is ugly.

    On Error GoTo notfound:
10        m_listSync.item (mode)
20        listIsSynced = True
          
notfound:
End Function
    
Public Sub syncListMode(mode As String)
10        m_session.sendLine "MODE " & m_name & " +" & mode
End Sub

Private Sub addListMode(mode As String, param As String)
          Dim listMode As New CListModeEntry
          
10        listMode.init mode, param
20        m_listModes.Add listMode
End Sub

Private Sub removeListMode(mode As String, param As String)
          Dim count As Long
          
10        For count = 1 To m_listModes.count
20            If m_listModes.item(count).mode = mode Then
30                If LCase$(m_listModes.item(count).param) = LCase$(param) Then
40                    m_listModes.Remove count
50                    Exit Sub
60                End If
70            End If
80        Next count
End Sub

Private Sub clearListMode(mode As String)
          Dim count As Long
          
10        For count = m_listModes.count To 1 Step -1
20            If m_listModes.item(count).mode = mode Then
30                m_listModes.Remove count
40            End If
50        Next count
          
60        setListSync mode, False
End Sub

Public Sub modeListEntry(mode As String, param As String)
10        If listIsSynced(mode) Then
20            clearListMode mode
30        End If
          
40        addListMode mode, param
End Sub

Public Sub endOfModeList(mode As String)
10        setListSync mode, True
20        RaiseEvent modeListSynced(mode)
End Sub

Private Sub addParamMode(mode As String, param As String)
          Dim paramMode As New CParamMode
          
10        paramMode.init mode, param
20        m_paramModes.Add paramMode
          
30        If mode = "k" Then
40            m_key = param
50        ElseIf mode = "l" Then
60            m_limit = Val(param)
70        End If
End Sub

Private Sub removeParamMode(mode As String)
          Dim count As Long
          
10        If mode = "k" Then
20            m_key = vbNullString
30        ElseIf mode = "l" Then
40            m_limit = 0
50        End If
          
60        For count = 1 To m_paramModes.count
70            If m_paramModes.item(count).mode = mode Then
80                m_paramModes.Remove count
90                Exit Sub
100           End If
110       Next count
End Sub

Public Function getModeParam(mode As String) As String
          Dim count As Long
          
10        For count = 1 To m_paramModes.count
20            If m_paramModes.item(count).mode = mode Then
30                getModeParam = m_paramModes.item(count).param
40                Exit Function
50            End If
60        Next count
End Function

Private Sub addMode(mode As String)
          Dim count As Long
          
10        For count = 1 To Len(m_modes)
20            If Mid$(m_modes, count, 1) = " " Then
30                Mid(m_modes, count, 1) = mode
40                Exit Sub
50            End If
60        Next count

70        m_modes = m_modes & mode
End Sub

Private Sub removeMode(mode As String)
          Dim count As Long
          
10        For count = 1 To Len(m_modes)
20            If Mid$(m_modes, count, 1) = mode Then
30                Mid(m_modes, count, 1) = " "
40                Exit Sub
50            End If
60        Next count
End Sub

Public Function hasMode(mode As String) As Boolean
          Dim count As Long
          
10        For count = 1 To Len(m_modes)
20            If Mid$(m_modes, count, 1) = mode Then
30                hasMode = True
40                Exit Function
50            End If
60        Next count
End Function

Public Sub modeChange(origin As COrigin, modes As String, paramString As String)
          Dim params() As String
          Dim currentParam As Long
          Dim maxParam As Long
          Dim count As Long
          Dim mode As String
          
          Dim modeOp As Boolean
          
10        modeOp = True
          
20        params = Split(paramString, " ")
30        maxParam = UBound(params)
          
40        For count = 1 To Len(modes)
50            mode = Mid$(modes, count, 1)
              
60            If mode = "+" Then
70                modeOp = True
80            ElseIf mode = "-" Then
90                modeOp = False
100           ElseIf m_session.isPrefixMode(mode) Then
110               If currentParam <= maxParam Then
120                   memberModeChange origin, mode, modeOp, params(currentParam)
130                   currentParam = currentParam + 1
140               End If
150           Else
                  Dim modeType As eChannelModeType
                  
160               modeType = m_session.getChannelModeType(mode)
                  
170               Select Case modeType
                      Case cmtList
180                       If currentParam <= maxParam Then
190                           If modeOp Then
200                               addListMode mode, params(currentParam)
210                           Else
220                               removeListMode mode, params(currentParam)
230                           End If
                              
240                           currentParam = currentParam + 1
250                       End If
260                   Case cmtParam
270                       If currentParam <= maxParam Then
280                           If modeOp Then
290                               addParamMode mode, params(currentParam)
300                           Else
310                               removeParamMode mode
320                           End If
                              
330                           currentParam = currentParam + 1
340                       End If
350                   Case cmtSetOnly
360                       If currentParam <= maxParam Or Not modeOp Then
370                           If modeOp Then
380                               addParamMode mode, params(currentParam)
390                           Else
400                               removeParamMode mode
410                           End If
420                       End If
430                   Case cmtNormal
440                       If modeOp Then
450                           addMode mode
460                       Else
470                           removeMode mode
480                       End If
490                   Case cmtUnknown
500               End Select
510           End If
520       Next count
          
530       m_windowText.addEvent "CHANNEL_MODE_CHANGE", makeStringArray(origin.nickname, m_name, modes, _
              paramString)
540       m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Private Sub memberModeChange(origin As COrigin, mode As String, modeOp As Boolean, member As String)
          Dim channelMember As CChannelMember
          
10        Set channelMember = findMember(member)
          
20        If Not channelMember Is Nothing Then
30            m_window.removeNicklistItem channelMember.nickname, _
                  m_session.getEffectivePrefix(channelMember.modes)
                  
40            If modeOp Then
50                channelMember.addMode mode
60            Else
70                channelMember.removeMode mode
80            End If
              
90            m_window.addNicklistItem channelMember.nickname, _
                  m_session.getEffectivePrefix(channelMember.modes)
100       End If
End Sub

Public Sub userMessage(origin As COrigin, text As String, Optional notice = False, Optional prefix _
    As String = vbNullString, Optional highlight As Boolean = False)

10        If notice Then
20            If LenB(prefix) <> 0 Then
30                m_windowText.addEventEx "WALLCHOP_NOTICE", getUserStyle(origin.nickname), _
                      origin.nickname, TVE_NONE, makeStringArray(prefix, m_name, text)
40            Else
50                m_windowText.addEventEx "CHANNEL_NOTICE", getUserStyle(origin.nickname), _
                      origin.nickname, TVE_NONE, makeStringArray(m_name, text)
60            End If
70        Else
80            If LenB(prefix) <> 0 Then
90                m_windowText.addEventEx "WALLCHOP_PRIVMSG", getUserStyle(origin.nickname), _
                      origin.nickname, TVE_NONE, makeStringArray(prefix, m_name, text)
100           Else
110               If highlight Then
120                   m_windowText.addEventEx "CHANNEL_PRIVMSG_HIGHLIGHT", getUserStyle(origin.nickname), _
                          origin.nickname, TVE_NONE, makeStringArray(text)
130                       m_session.client.switchbar.tabActivity Me.window.switchbartab, tasHighlight
140               Else
150                   m_windowText.addEventEx "CHANNEL_PRIVMSG", getUserStyle(origin.nickname), _
                          origin.nickname, TVE_NONE, makeStringArray(text)
160               End If
170           End If
180       End If
          
190       m_session.client.switchbar.tabActivity Me.window.switchbartab, tasMessage
End Sub

Public Sub userEmote(origin As COrigin, text As String, Optional prefix As String = vbNullString, Optional highlight As Boolean = False)
10        If LenB(prefix) <> 0 Then
20            m_windowText.addEventEx "WALLCHOP_EMOTE", getUserStyle(origin.nickname), origin.nickname, _
                  TVE_NONE, makeStringArray(prefix, m_name, text)
30        Else
40            If highlight Then
50                m_windowText.addEventEx "CHANNEL_EMOTE_HIGHLIGHT", getUserStyle(origin.nickname), origin.nickname, _
                      TVE_NONE, makeStringArray(text)
60            Else
70                m_windowText.addEventEx "CHANNEL_EMOTE", getUserStyle(origin.nickname), origin.nickname, _
                      TVE_NONE, makeStringArray(text)
80            End If
90        End If
          
100       m_session.client.switchbar.tabActivity Me.window.switchbartab, tasMessage
End Sub

Public Sub userCtcp(origin As COrigin, command As String, param As String, Optional prefix As _
    String = vbNullString)
          
10        If LenB(prefix) <> 0 Then
20            m_windowText.addEventEx "WALLCHOP_CTCP", getUserStyle(origin.nickname), origin.nickname, _
                  TVE_NONE, makeStringArray(prefix, m_name, command, param)
30        Else
40            m_windowText.addEventEx "CHANNEL_CTCP", getUserStyle(origin.nickname), origin.nickname, _
                  TVE_NONE, makeStringArray(m_name, command, param)
50        End If
          
60        m_session.client.switchbar.tabActivity Me.window.switchbartab, tasMessage
End Sub

Private Function getUserStyle(nickname As String) As CUserStyle
          Dim member As CChannelMember
          
10        Set member = findMember(nickname)
          
20        If member Is Nothing Then
30            Exit Function
40        End If
          
50        Set getUserStyle = m_session.getUserStyle(member.nickname, m_session.getEffectivePrefix(member.modes))
End Function

Public Sub processNamesList(names As String)
10        If m_membersListSynced Then
20            clearMembers
30        End If

          Dim count As Integer
          Dim char As String
          
          Dim nick As String
          Dim modes As String
          Dim mode As String
          
40        If right(names, 1) <> " " Then
50            names = names & " "
60        End If
          
70        For count = 1 To Len(names)
80            char = Mid$(names, count, 1)
              
90            If AscW(char) = 32 Then
100               If LenB(nick) <> 0 Then
110                   If findMember(nick) Is Nothing Then
120                       addChannelMember nick, modes
130                   End If
                      
140                   nick = vbNullString
150                   modes = vbNullString
160               End If
170           Else
180               If LenB(nick) <> 0 Then
190                   nick = nick & char
200               Else
210                   mode = m_session.getPrefixModeBySymbol(char)
                      
220                   If LenB(mode) <> 0 Then
230                       modes = modes & mode
240                   Else
250                       nick = nick & char
260                   End If
270               End If
280           End If
290       Next count
End Sub

Public Sub endOfNames()
10        m_membersListSynced = True
End Sub

Public Function isIn(nickname As String) As Boolean
          Dim member As CChannelMember
          
10        Set member = findMember(nickname)
          
20        isIn = Not member Is Nothing
End Function

Private Function addChannelMember(nickname As String, Optional modes As String = vbNullString)
          Dim channelMember As New CChannelMember
          
10        channelMember.nickname = nickname
20        channelMember.modes = modes

30        m_members.Add channelMember, LCase(nickname)
40        m_window.addNicklistItem nickname, m_session.getEffectivePrefix(modes)
End Function

Private Sub removeChannelMember(nickname As String)
          Dim channelMember As CChannelMember
          
10        Set channelMember = findMember(nickname)
          
20        If Not channelMember Is Nothing Then
30            m_window.removeNicklistItem nickname, m_session.getEffectivePrefix(channelMember.modes)
40            m_members.Remove LCase$(channelMember.nickname)
50        End If
End Sub

Private Sub clearMembers()
          Dim channelMember As CChannelMember
          
10        For Each channelMember In m_members
20            m_members.Remove LCase$(channelMember.nickname)
30        Next channelMember
          
40        m_window.clearNicklist
50        m_membersListSynced = False
End Sub

Private Function findMember(nickname As String) As CChannelMember
10        On Error Resume Next
20        Set findMember = m_members.item(LCase(nickname))
End Function

Public Sub textEntered(text As String)
10        m_session.sendLine "PRIVMSG " & m_name & " :" & text
20        m_windowText.addEventEx "ME_PRIVMSG", getUserStyle(m_session.currentNickname), _
              m_session.currentNickname, TVE_NONE, makeStringArray(text)
End Sub

Public Sub sendEmote(text As String)
10        m_session.sendLine "PRIVMSG " & m_name & " :" & Chr$(1) & "ACTION " & text & Chr$(1)
20        m_windowText.addEventEx "CHANNEL_EMOTE", getUserStyle(m_session.currentNickname), _
              m_session.currentNickname, TVE_NONE, makeStringArray(text)
End Sub

Private Sub Class_Terminate()
10        debugLog "CChannel terminated"
End Sub

Private Sub m_rejoinTimer_Timer()
10        m_rejoinTimer.enabled = False
20        m_session.rejoinChannel Me
End Sub
