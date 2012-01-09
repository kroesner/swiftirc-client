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

Implements IOptionsChangeListener

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
    meIsIn = m_isin
End Property

Public Property Get session() As CSession
    Set session = m_session
End Property

Public Property Get topic() As String
    topic = m_topic
End Property

Public Property Get limit() As Long
    limit = m_limit
End Property

Public Property Get key() As String
    key = m_key
End Property

Public Sub init(name As String, window As ctlWindowChannel, session As CSession)
    m_name = name
    Set m_window = window
    Set m_windowText = window
    Set m_session = session
    
    Set m_rejoinTimer = m_session.client.getTimer
    
    refreshLogging
    registerForOptionsChanges Me
End Sub

Public Sub deInit()
    m_window.deInit

    m_session.client.releaseTimer m_rejoinTimer

    Set m_window = Nothing
    Set m_windowText = Nothing
    Set m_session = Nothing
    
    unregisterForOptionsChanges Me
End Sub

Public Property Get IText() As ITextWindow
    Set IText = m_window
End Property

Public Property Get window() As ctlWindowChannel
    Set window = m_window
End Property

Public Property Get name() As String
    name = m_name
End Property

Public Property Get hopping() As Boolean
    hopping = m_hopping
End Property

Public Property Let hopping(newValue As Boolean)
    m_hopping = newValue
End Property

Public Sub disconnected()
    clearMembers
    m_windowText.addEvent "DISCONNECTED", makeStringArray("")
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
    m_isin = False
End Sub

Public Sub reJoining()
    m_windowText.addEvent "ME_REJOINING", makeStringArray(m_name)
    clearMembers
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
    m_isin = False
End Sub

Public Sub reJoined()
    m_windowText.addEvent "ME_REJOINED", makeStringArray(m_name)
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
    m_isin = True
End Sub

Public Sub meJoined()
    m_windowText.addEvent "ME_JOIN", makeStringArray(m_name)
    m_session.sendLine "MODE " & m_name
    m_isin = True
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub hop(Optional reason As String = "Hopping")
    m_hopping = True
    m_session.sendLine "PART " & m_name & " :" & reason
End Sub

Public Sub userJoined(origin As COrigin)
    If Not findMember(origin.nickname) Is Nothing Then
        Exit Sub
    End If
    
    addChannelMember origin.nickname
    m_windowText.addEvent "USER_JOIN", makeStringArray(origin.nickname, m_name, origin.username, _
        origin.hostname)
        
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userParted(origin As COrigin, Optional reason As String = vbNullString)
    removeChannelMember origin.nickname
    
    If LenB(reason) = 0 Then
        m_windowText.addEvent "USER_PART", makeStringArray(origin.nickname, m_name, origin.username, _
            origin.hostname)
    Else
        m_windowText.addEvent "USER_PART_REASON", makeStringArray(origin.nickname, m_name, reason, _
            origin.username, origin.hostname)
    End If
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userQuit(origin As COrigin, Optional reason As String = vbNullString)
    If LenB(reason) = 0 Then
        m_windowText.addEventEx "USER_QUIT", getUserStyle(origin.nickname), origin.nickname, _
            TVE_NONE, makeStringArray("")
    Else
        m_windowText.addEventEx "USER_QUIT_REASON", getUserStyle(origin.nickname), origin.nickname, _
            TVE_NONE, makeStringArray(reason)
    End If
    
    removeChannelMember origin.nickname
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userNickChange(origin As COrigin, newNick As String)
    Dim member As CChannelMember
    
    Set member = findMember(origin.nickname)
    
    If member Is Nothing Then
        Exit Sub
    End If
    
    m_window.removeNicklistItem member.nickname, m_session.getEffectivePrefix(member.modes)
    m_members.Remove LCase$(member.nickname)
    member.nickname = newNick
    m_members.Add member, LCase$(newNick)
    m_window.addNicklistItem member.nickname, m_session.getEffectivePrefix(member.modes)
    
    m_windowText.addEvent "USER_NICK_CHANGE", makeStringArray(origin.nickname, newNick)
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub userKicked(origin As COrigin, target As String, reason As String)
    removeChannelMember target
    m_windowText.addEvent "USER_KICKED", makeStringArray(origin.nickname, m_name, target, reason)
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub meKicked(origin As COrigin, reason As String)
    clearMembers
    m_isin = False
    m_windowText.addEvent "ME_KICKED", makeStringArray(m_name, origin.nickname, reason)
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
    
    If settings.autoRejoinOnKick Then
        m_windowText.addEvent "ME_REJOIN_DELAY", makeStringArray(m_name, "2 seconds")
        m_rejoinTimer.Interval = 2000
        m_rejoinTimer.enabled = True
    End If
End Sub

Private Sub addTopicHistory(topic As String)
    Dim tempTopic As Variant
    
    For Each tempTopic In m_topicHistory
        If StrComp(tempTopic, topic, vbBinaryCompare) = 0 Then
            Exit Sub
        End If
    Next tempTopic
    
    m_topicHistory.Add topic
End Sub

Public Function getTopicHistory(index As Long) As String
    getTopicHistory = m_topicHistory.item(index)
End Function

Public Function getTopicHistoryCount() As Long
    getTopicHistoryCount = m_topicHistory.count
End Function

Public Sub topicIs(topic As String)
    m_windowText.addEvent "CHANNEL_TOPICIS", makeStringArray(topic)
    m_topic = topic
    
    addTopicHistory topic
End Sub

Public Sub topicWhoTime(who As String, time As Long)
    m_windowText.addEvent "CHANNEL_TOPICWHOTIME", makeStringArray(who, formatTime(time))
End Sub

Public Sub topicChange(origin As COrigin, topic As String)
    m_windowText.addEvent "CHANNEL_TOPICCHANGE", makeStringArray(origin.nickname, topic)
    m_topic = topic
    
    addTopicHistory topic
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Public Sub modeIs(modes As String, Optional paramString As String = vbNullString)
    m_paramModes.clear
    m_modes = vbNullString

    Dim count As Long
    Dim params() As String
    Dim mode As String
    
    Dim paramIndex As Long
    Dim paramCount As Long
    
    Dim modeType As eChannelModeType
            
    params = Split(paramString, " ")
    paramCount = UBound(params) + 1
    
    For count = 1 To Len(modes)
        mode = Mid$(modes, count, 1)
        modeType = m_session.getChannelModeType(mode)
        
        Select Case modeType
            Case cmtParam
                If paramIndex < paramCount Then
                    addParamMode mode, params(paramIndex)
                    paramIndex = paramIndex + 1
                End If
            Case cmtSetOnly
                If paramIndex < paramCount Then
                    addParamMode mode, params(paramIndex)
                    paramIndex = paramIndex + 1
                End If
            Case cmtNormal
                addMode mode
        End Select
    Next count
End Sub
    
Public Sub getModeList(mode As String, list As cArrayList)
    Dim count As Long
    
    For count = 1 To m_listModes.count
        If m_listModes.item(count).mode = mode Then
            list.Add m_listModes.item(count)
        End If
    Next count
End Sub
    
Private Sub setListSync(mode As String, op As Boolean)
    On Error Resume Next
    
    If op Then
        m_listSync.Add CVar(True), mode
    Else
        m_listSync.Remove mode
    End If
End Sub

Public Function listIsSynced(mode As String) As Boolean
    'This is ugly.

    On Error GoTo notfound:
    m_listSync.item (mode)
    listIsSynced = True
    
notfound:
End Function
    
Public Sub syncListMode(mode As String)
    m_session.sendLine "MODE " & m_name & " +" & mode
End Sub

Private Sub addListMode(mode As String, param As String)
    Dim listMode As New CListModeEntry
    
    listMode.init mode, param
    m_listModes.Add listMode
End Sub

Private Sub removeListMode(mode As String, param As String)
    Dim count As Long
    
    For count = 1 To m_listModes.count
        If m_listModes.item(count).mode = mode Then
            If LCase$(m_listModes.item(count).param) = LCase$(param) Then
                m_listModes.Remove count
                Exit Sub
            End If
        End If
    Next count
End Sub

Private Sub clearListMode(mode As String)
    Dim count As Long
    
    For count = m_listModes.count To 1 Step -1
        If m_listModes.item(count).mode = mode Then
            m_listModes.Remove count
        End If
    Next count
    
    setListSync mode, False
End Sub

Public Sub modeListEntry(mode As String, param As String)
    If listIsSynced(mode) Then
        clearListMode mode
    End If
    
    addListMode mode, param
End Sub

Public Sub endOfModeList(mode As String)
    setListSync mode, True
    RaiseEvent modeListSynced(mode)
End Sub

Private Sub addParamMode(mode As String, param As String)
    Dim paramMode As New CParamMode
    
    paramMode.init mode, param
    m_paramModes.Add paramMode
    
    If mode = "k" Then
        m_key = param
    ElseIf mode = "l" Then
        m_limit = Val(param)
    End If
End Sub

Private Sub removeParamMode(mode As String)
    Dim count As Long
    
    If mode = "k" Then
        m_key = vbNullString
    ElseIf mode = "l" Then
        m_limit = 0
    End If
    
    For count = 1 To m_paramModes.count
        If m_paramModes.item(count).mode = mode Then
            m_paramModes.Remove count
            Exit Sub
        End If
    Next count
End Sub

Public Function getModeParam(mode As String) As String
    Dim count As Long
    
    For count = 1 To m_paramModes.count
        If m_paramModes.item(count).mode = mode Then
            getModeParam = m_paramModes.item(count).param
            Exit Function
        End If
    Next count
End Function

Private Sub addMode(mode As String)
    Dim count As Long
    
    For count = 1 To Len(m_modes)
        If Mid$(m_modes, count, 1) = " " Then
            Mid(m_modes, count, 1) = mode
            Exit Sub
        End If
    Next count

    m_modes = m_modes & mode
End Sub

Private Sub removeMode(mode As String)
    Dim count As Long
    
    For count = 1 To Len(m_modes)
        If Mid$(m_modes, count, 1) = mode Then
            Mid(m_modes, count, 1) = " "
            Exit Sub
        End If
    Next count
End Sub

Public Function hasMode(mode As String) As Boolean
    Dim count As Long
    
    For count = 1 To Len(m_modes)
        If Mid$(m_modes, count, 1) = mode Then
            hasMode = True
            Exit Function
        End If
    Next count
End Function

Public Sub modeChange(origin As COrigin, modes As String, paramString As String)
    Dim params() As String
    Dim currentParam As Long
    Dim maxParam As Long
    Dim count As Long
    Dim mode As String
    
    Dim modeOp As Boolean
    
    modeOp = True
    
    params = Split(paramString, " ")
    maxParam = UBound(params)
    
    For count = 1 To Len(modes)
        mode = Mid$(modes, count, 1)
        
        If mode = "+" Then
            modeOp = True
        ElseIf mode = "-" Then
            modeOp = False
        ElseIf m_session.isPrefixMode(mode) Then
            If currentParam <= maxParam Then
                memberModeChange origin, mode, modeOp, params(currentParam)
                currentParam = currentParam + 1
            End If
        Else
            Dim modeType As eChannelModeType
            
            modeType = m_session.getChannelModeType(mode)
            
            Select Case modeType
                Case cmtList
                    If currentParam <= maxParam Then
                        If modeOp Then
                            addListMode mode, params(currentParam)
                        Else
                            removeListMode mode, params(currentParam)
                        End If
                        
                        currentParam = currentParam + 1
                    End If
                Case cmtParam
                    If currentParam <= maxParam Then
                        If modeOp Then
                            addParamMode mode, params(currentParam)
                        Else
                            removeParamMode mode
                        End If
                        
                        currentParam = currentParam + 1
                    End If
                Case cmtSetOnly
                    If currentParam <= maxParam Or Not modeOp Then
                        If modeOp Then
                            addParamMode mode, params(currentParam)
                        Else
                            removeParamMode mode
                        End If
                    End If
                Case cmtNormal
                    If modeOp Then
                        addMode mode
                    Else
                        removeMode mode
                    End If
                Case cmtUnknown
            End Select
        End If
    Next count
    
    m_windowText.addEvent "CHANNEL_MODE_CHANGE", makeStringArray(origin.nickname, m_name, modes, _
        paramString)
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasEvent
End Sub

Private Sub memberModeChange(origin As COrigin, mode As String, modeOp As Boolean, member As String)
    Dim channelMember As CChannelMember
    
    Set channelMember = findMember(member)
    
    If Not channelMember Is Nothing Then
        m_window.removeNicklistItem channelMember.nickname, _
            m_session.getEffectivePrefix(channelMember.modes)
            
        If modeOp Then
            channelMember.addMode mode
        Else
            channelMember.removeMode mode
        End If
        
        m_window.addNicklistItem channelMember.nickname, _
            m_session.getEffectivePrefix(channelMember.modes)
    End If
End Sub

Public Sub userMessage(origin As COrigin, text As String, Optional notice = False, Optional prefix _
    As String = vbNullString, Optional highlight As Boolean = False)

    If notice Then
        If LenB(prefix) <> 0 Then
            m_windowText.addEventEx "WALLCHOP_NOTICE", getUserStyle(origin.nickname), _
                origin.nickname, TVE_NONE, makeStringArray(prefix, m_name, text)
        Else
            m_windowText.addEventEx "CHANNEL_NOTICE", getUserStyle(origin.nickname), _
                origin.nickname, TVE_NONE, makeStringArray(m_name, text)
        End If
    Else
        If LenB(prefix) <> 0 Then
            m_windowText.addEventEx "WALLCHOP_PRIVMSG", getUserStyle(origin.nickname), _
                origin.nickname, TVE_NONE, makeStringArray(prefix, m_name, text)
        Else
            If highlight Then
                m_windowText.addEventEx "CHANNEL_PRIVMSG_HIGHLIGHT", getUserStyle(origin.nickname), _
                    origin.nickname, TVE_NONE, makeStringArray(text)
                    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasHighlight
            Else
                m_windowText.addEventEx "CHANNEL_PRIVMSG", getUserStyle(origin.nickname), _
                    origin.nickname, TVE_NONE, makeStringArray(text)
            End If
        End If
    End If
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasMessage
End Sub

Public Sub userEmote(origin As COrigin, text As String, Optional prefix As String = vbNullString, Optional highlight As Boolean = False)
    If LenB(prefix) <> 0 Then
        m_windowText.addEventEx "WALLCHOP_EMOTE", getUserStyle(origin.nickname), origin.nickname, _
            TVE_NONE, makeStringArray(prefix, m_name, text)
    Else
        If highlight Then
            m_windowText.addEventEx "CHANNEL_EMOTE_HIGHLIGHT", getUserStyle(origin.nickname), origin.nickname, _
                TVE_NONE, makeStringArray(text)
        Else
            m_windowText.addEventEx "CHANNEL_EMOTE", getUserStyle(origin.nickname), origin.nickname, _
                TVE_NONE, makeStringArray(text)
        End If
    End If
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasMessage
End Sub

Public Sub userCtcp(origin As COrigin, command As String, param As String, Optional prefix As _
    String = vbNullString)
    
    If LenB(prefix) <> 0 Then
        m_windowText.addEventEx "WALLCHOP_CTCP", getUserStyle(origin.nickname), origin.nickname, _
            TVE_NONE, makeStringArray(prefix, m_name, command, param)
    Else
        m_windowText.addEventEx "CHANNEL_CTCP", getUserStyle(origin.nickname), origin.nickname, _
            TVE_NONE, makeStringArray(m_name, command, param)
    End If
    
    m_session.client.switchbar.tabActivity Me.window.switchbartab, tasMessage
End Sub

Private Function getUserStyle(nickname As String) As CUserStyle
    Dim member As CChannelMember
    
    Set member = findMember(nickname)
    
    If member Is Nothing Then
        Exit Function
    End If
    
    Set getUserStyle = m_session.getUserStyle(member.nickname, m_session.getEffectivePrefix(member.modes))
End Function

Public Sub processNamesList(names As String)
    If m_membersListSynced Then
        clearMembers
    End If

    Dim count As Integer
    Dim char As String
    
    Dim nick As String
    Dim modes As String
    Dim mode As String
    
    If right(names, 1) <> " " Then
        names = names & " "
    End If
    
    For count = 1 To Len(names)
        char = Mid$(names, count, 1)
        
        If AscW(char) = 32 Then
            If LenB(nick) <> 0 Then
                If findMember(nick) Is Nothing Then
                    addChannelMember nick, modes
                End If
                
                nick = vbNullString
                modes = vbNullString
            End If
        Else
            If LenB(nick) <> 0 Then
                nick = nick & char
            Else
                mode = m_session.getPrefixModeBySymbol(char)
                
                If LenB(mode) <> 0 Then
                    modes = modes & mode
                Else
                    nick = nick & char
                End If
            End If
        End If
    Next count
End Sub

Public Sub endOfNames()
    m_membersListSynced = True
End Sub

Public Function isIn(nickname As String) As Boolean
    Dim member As CChannelMember
    
    Set member = findMember(nickname)
    
    isIn = Not member Is Nothing
End Function

Private Function addChannelMember(nickname As String, Optional modes As String = vbNullString)
    Dim channelMember As New CChannelMember
    
    channelMember.nickname = nickname
    channelMember.modes = modes

    m_members.Add channelMember, LCase(nickname)
    m_window.addNicklistItem nickname, m_session.getEffectivePrefix(modes)
End Function

Private Sub removeChannelMember(nickname As String)
    Dim channelMember As CChannelMember
    
    Set channelMember = findMember(nickname)
    
    If Not channelMember Is Nothing Then
        m_window.removeNicklistItem nickname, m_session.getEffectivePrefix(channelMember.modes)
        m_members.Remove LCase$(channelMember.nickname)
    End If
End Sub

Private Sub clearMembers()
    Dim channelMember As CChannelMember
    
    For Each channelMember In m_members
        m_members.Remove LCase$(channelMember.nickname)
    Next channelMember
    
    m_window.clearNicklist
    m_membersListSynced = False
End Sub

Private Function findMember(nickname As String) As CChannelMember
    On Error Resume Next
    Set findMember = m_members.item(LCase(nickname))
End Function

Public Sub textEntered(text As String)
    m_session.sendLine "PRIVMSG " & m_name & " :" & text
    m_windowText.addEventEx "ME_PRIVMSG", getUserStyle(m_session.currentNickname), _
        m_session.currentNickname, TVE_NONE, makeStringArray(text)
End Sub

Public Sub sendEmote(text As String)
    m_session.sendLine "PRIVMSG " & m_name & " :" & Chr$(1) & "ACTION " & text & Chr$(1)
    m_windowText.addEventEx "CHANNEL_EMOTE", getUserStyle(m_session.currentNickname), _
        m_session.currentNickname, TVE_NONE, makeStringArray(text)
End Sub

Private Sub IOptionsChangeListener_optionsChanged()
    refreshLogging
End Sub

Private Sub refreshLogging()
    m_windowText.textview.logName = combinePath(m_session.baseLogPath, sanitizeFilename(m_name))
    m_windowText.textview.enableLogging = shouldLog
End Sub

Private Function shouldLog() As Boolean
    shouldLog = settings.setting("enableLogging", estBoolean) And settings.setting("logChannel", estBoolean)
End Function


Private Sub m_rejoinTimer_Timer()
    m_rejoinTimer.enabled = False
    m_session.rejoinChannel Me
End Sub
