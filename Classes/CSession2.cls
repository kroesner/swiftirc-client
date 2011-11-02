VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const RECONNECT_DELAY As Long = 10000

Private m_serverProfile As CServerProfile
Private m_client As swiftIRc.SwiftIrcClient

Private WithEvents m_statusWindow As ctlWindowStatus
Attribute m_statusWindow.VB_VarHelpID = -1
Private m_statusWindowText As ITextWindow
Private m_debugWindow As ctlWindowGenericText
Private m_debugWindowText As ITextWindow

Private WithEvents m_reconnectTimer As VB.timer
Attribute m_reconnectTimer.VB_VarHelpID = -1

Private WithEvents m_socket As CSocket
Attribute m_socket.VB_VarHelpID = -1

Private m_parsers As New Collection
Private m_inputParsers As New Collection

Private m_serverHost As String
Private m_serverPort As Long
Private m_serverPassword As String

Private m_serverName As String
Private m_networkName As String

Private m_currentNickname As String

Private m_primaryNickname As String
Private m_backupNickname As String
Private m_username As String
Private m_realName As String

Private m_expectingDisconnect As Boolean
Private m_connected As Boolean
Private m_registered As Boolean
Private m_triedBackupNickname As Boolean

Private m_currentLine As String

Private m_channels As New Collection
Private m_queries As New Collection

Private m_iSupport As New Collection
Private m_maxModes As Byte
Private m_prefixModes As New cArrayList
Private m_channelModes As New cArrayList

'Channel tabbing/completion
Private m_lastChannelTab As String
Private m_lastChannelTabIndex As Long
Private m_lastChannelTabMatch As String

Public Property Get serverProfile() As CServerProfile
    Set serverProfile = m_serverProfile
End Property

Public Property Let serverProfile(newValue As CServerProfile)
    Set m_serverProfile = newValue
    
    m_primaryNickname = m_serverProfile.primaryNickname
    m_backupNickname = m_serverProfile.backupNickname
    m_username = m_serverProfile.primaryNickname
    
    m_currentNickname = m_primaryNickname
    
    If LenB(m_serverProfile.realName) <> 0 Then
        m_realName = m_serverProfile.realName
    Else
        m_realName = m_serverProfile.primaryNickname
    End If
    
    m_serverHost = m_serverProfile.hostname
    
    If m_serverProfile.port <> 0 Then
        m_serverPort = m_serverProfile.port
    Else
        m_serverPort = 6667
    End If
    
    m_serverPassword = m_serverProfile.serverPassword
    
    m_statusWindow.switchbarTab.caption = m_serverHost
    m_client.redrawSwitchbarTab m_statusWindow.switchbarTab
End Property

Public Property Get client() As swiftIRc.SwiftIrcClient
    Set client = m_client
End Property

Public Property Let client(newValue As swiftIRc.SwiftIrcClient)
    Set m_client = newValue
End Property

Public Property Get statusWindow() As ctlWindowStatus
    Set statusWindow = m_statusWindow
End Property

Public Property Let statusWindow(newValue As ctlWindowStatus)
    Set m_statusWindow = newValue
    Set m_statusWindowText = newValue
End Property

Public Property Get serverHost() As String
    serverHost = m_serverHost
End Property

Public Property Let serverHost(newValue As String)
    m_serverHost = newValue
End Property

Public Property Get serverPort() As Long
    serverPort = m_serverPort
End Property

Public Property Let serverPort(newValue As Long)
    m_serverPort = newValue
End Property

Public Property Get serverPassword() As String
    serverPassword = m_serverPassword
End Property

Public Property Let serverPassword(newValue As String)
    m_serverPassword = newValue
End Property

Public Property Get currentNickname() As String
    currentNickname = m_currentNickname
End Property

Public Property Get primaryNickname() As String
    primaryNickname = m_primaryNickname
End Property

Public Property Let primaryNickname(newValue As String)
    m_primaryNickname = newValue
    m_currentNickname = newValue
    m_username = newValue
End Property

Public Property Get backupNickname() As String
    backupNickname = m_backupNickname
End Property

Public Property Let backupNickname(newValue As String)
    m_backupNickname = newValue
End Property

Public Property Get username() As String
    username = m_username
End Property

Public Property Let username(newValue As String)
    m_username = newValue
End Property

Public Property Get realName() As String
    realName = m_realName
End Property

Public Property Let realName(newValue As String)
    m_realName = newValue
End Property

Public Property Get networkName() As String
    networkName = m_networkName
End Property

Public Property Get connected() As Boolean
    connected = m_connected
End Property

Public Sub connect()
    m_reconnectTimer.Enabled = False

    If m_connected Then
        disconnect
    End If

    m_statusWindowText.addEvent "CONNECTING", makeStringArray(m_serverHost, m_serverPort)
    
    If m_serverPort = 0 Then
        m_serverPort = 6667
    End If
    
    If LenB(m_username) = 0 Then
        m_username = m_primaryNickname
    End If
    
    If LenB(m_realName) = 0 Then
        m_realName = m_primaryNickname
    End If
    
    m_socket.connect m_serverHost, m_serverPort
End Sub

Public Sub disconnect()
    m_expectingDisconnect = True
    'm_socket.sendLine "QUIT"
    m_socket.closeSocket
    disconnected
End Sub

Private Sub login()
    m_triedBackupNickname = False
    
    m_statusWindowText.addEvent "CONNECTED", makeStringArray(m_serverHost)

    If LenB(m_serverPassword) <> 0 Then
        m_socket.sendLine "PASS :" & m_serverPassword
    End If

    m_socket.sendLine "USER " & m_username & " * 0 :" & m_realName
    m_socket.sendLine "NICK " & m_currentNickname
End Sub

Private Sub disconnected()
    m_statusWindowText.addEvent "DISCONNECTED", makeStringArray("")
    
    Dim channel As CChannel
    Dim textWnd As ITextWindow
    
    For Each channel In m_channels
        channel.disconnected
    Next channel
    
    Do While m_iSupport.count <> 0
        m_iSupport.Remove 1
    Loop
    
    m_networkName = vbNullString
    
    m_prefixModes.Clear
    m_channelModes.Clear
    m_maxModes = 0
    
    m_registered = False
    m_connected = False
    
    If Not m_expectingDisconnect Then
        If Not m_serverProfile Is Nothing Then
            If m_serverProfile.enableReconnect Then
                m_statusWindowText.addEvent "RECONNECTING_IN", makeStringArray(RECONNECT_DELAY / 1000)
                m_reconnectTimer.Enabled = True
            End If
        End If
    End If
    
    m_expectingDisconnect = False
End Sub

Private Sub connectFailed()
    If Not m_serverProfile Is Nothing Then
        If m_serverProfile.enableConnectRetry Then
            m_statusWindowText.addEvent "RECONNECTING_IN", makeStringArray(RECONNECT_DELAY / 1000)
            m_reconnectTimer.Enabled = True
        End If
    End If
End Sub

Private Sub m_reconnectTimer_Timer()
    connect
End Sub

Private Sub parse(buffer As String)
    Dim count As Integer
    Dim char As Integer
    Dim last As Integer
    
    For count = 1 To Len(buffer)
        char = AscW(Mid$(buffer, count, 1))
        
        If char = 10 Or char = 13 Then
            If count < Len(buffer) Then
                char = AscW(Mid$(buffer, count + 1, 1))
                
                If char = 10 Or char = 13 Then
                    count = count + 1
                End If
            End If
            
            parseLine
            m_currentLine = vbNullString
        Else
            m_currentLine = m_currentLine & ChrW$(char)
        End If
    Next count
End Sub

Private Sub parseLine()
    Dim origin As String
    Dim command As String
    
    Dim line As String
    Dim fullLine As String
    
    line = m_currentLine
    fullLine = m_currentLine
    m_currentLine = vbNullString
    
    extractIrcHeader line, origin, command
    
    If LenB(command) = 0 Then
        Exit Sub
    End If
    
    If Not m_debugWindowText Is Nothing Then
        m_debugWindowText.addText "<- " & fullLine
    End If
    
    Dim parser As CParser
    
    Set parser = findParser(command)
    
    If parser Is Nothing Then
        Exit Sub
    End If
    
    Dim params() As String
    
    extractIrcParams line, parser.maxParams, params
    
    If UBound(params) + 1 < parser.minParams Then
        Exit Sub
    End If
    
    Dim parsedOrigin As New COrigin
    parsedOrigin.parseOrigin origin
    
    CallByName Me, parser.func, VbMethod, parsedOrigin, command, UBound(params) + 1, params
End Sub

Private Sub extractIrcHeader(ByRef line As String, ByRef origin As String, ByRef command As String)
    Dim currentParam As Byte

    currentParam = 1
    
    Do While LenB(line) <> 0
        If Mid$(line, 1, 1) = " " Then
            If LenB(command) <> 0 Then
                line = Mid$(line, 2)
                
                If LenB(origin) = 0 Then
                    origin = m_serverName
                End If
                
                Exit Sub
            Else
                currentParam = 1
            End If
        ElseIf Mid$(line, 1, 1) = ":" Then
            If LenB(command) = 0 Then
                currentParam = 0
            End If
        Else
            If currentParam = 0 Then
                origin = origin & Mid$(line, 1, 1)
            Else
                command = command & Mid$(line, 1, 1)
            End If
        End If
        
        line = Mid$(line, 2)
    Loop
End Sub

Private Sub extractIrcParams(ByRef line As String, maxParams As Byte, ByRef params() As String)
    Dim temp As String
    Dim paramCount As Byte
    
    Do While LenB(line) <> 0
        If left$(line, 2) = " :" Then
            If LenB(temp) <> 0 Then
                paramCount = paramCount + 1
                ReDim Preserve params(paramCount - 1)
                params(paramCount - 1) = temp
            End If
        
            line = RTrim(Mid$(line, 3))
            paramCount = paramCount + 1
            ReDim Preserve params(paramCount - 1)
            params(paramCount - 1) = line
            
            Exit Sub
        ElseIf left$(line, 1) = " " Then
            If LenB(temp) <> 0 Then
                paramCount = paramCount + 1
                ReDim Preserve params(paramCount - 1)
                params(paramCount - 1) = temp
                temp = vbNullString
            End If
        Else
            temp = temp & left$(line, 1)
        End If
        
        If paramCount + 1 >= maxParams Then
            Do While left$(line, 1) = " "
                line = Mid$(line, 2)
            Loop
            
            If LenB(line) = 0 Then
                Exit Sub
            End If
            
            If left$(line, 1) = ":" Then
                line = Mid$(line, 2)
            End If
            
            If LenB(line) = 0 Then
                Exit Sub
            End If
            
            paramCount = paramCount + 1
            ReDim Preserve params(paramCount - 1)
            params(paramCount - 1) = RTrim(line)
            Exit Sub
        End If
        
        line = Mid$(line, 2)
    Loop
    
    If LenB(temp) <> 0 Then
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount - 1)
        params(paramCount - 1) = temp
    End If
End Sub

Private Function addParser(ByVal name As String, ByVal minParams As Byte, _
    ByVal maxParams As Byte, ByVal func As String)
    
    Dim parser As CParser
    
    Set parser = New CParser
    
    parser.name = name
    parser.minParams = minParams
    parser.maxParams = maxParams
    parser.func = func
    
    m_parsers.Add parser, LCase$(name)
End Function

Private Function findParser(name As String) As CParser
    On Error Resume Next
    Set findParser = m_parsers.item(LCase$(name))
End Function
    
Private Sub initParsers()
    addParser "PING", 1, 1, "parserPing"
    addParser "ERROR", 1, 1, "parserError"

    addParser "KILL", 2, 2, "parserKill"

    addParser "001", 2, 2, "parserWelcome"
    addParser "005", 2, 2, "parserISupport"
    addParser "NOTICE", 2, 2, "parserNotice"
    addParser "JOIN", 1, 1, "parserJoin"
    addParser "PART", 1, 2, "parserPart"
    addParser "KICK", 3, 3, "parserKick"
    
    addParser "TOPIC", 2, 2, "parserTopic"
    
    addParser RPL_TOPIC, 3, 3, "parserTopicIs"
    addParser RPL_TOPICWHOTIME, 4, 4, "parserTopicWhoTime"
    
    addParser "353", 4, 4, "parserNames"
    addParser ERR_NICKNAMEINUSE, 3, 3, "parserNicknameInUse"
    addParser "PRIVMSG", 2, 2, "parserPrivmsg"
    addParser "MODE", 2, 3, "parserMode"
    
    addParser RPL_CHANNELMODEIS, 3, 4, "parserChannelModeIs"
    
    addParser RPL_BANLIST, 3, 5, "parserBanList"
    addParser RPL_ENDOFBANLIST, 3, 3, "parserEndOfBanList"
    
    addParser RPL_EXLIST, 3, 5, "parserExList"
    addParser RPL_ENDOFEXLIST, 3, 3, "parserEndOfExList"
    
    addParser RPL_INVEXLIST, 3, 5, "parserInvexList"
    addParser RPL_ENDOFINVEXLIST, 3, 3, "parserEndOfInvexList"
    
    addParser RPL_ALIST, 3, 3, "parserAlist"
    addParser RPL_ENDOFALIST, 3, 3, "parserEndOfAlist"
    
    addParser RPL_QLIST, 3, 3, "parserQList"
    addParser RPL_ENDOFQLIST, 3, 3, "parserEndOfQList"
End Sub

Public Sub parserPing(origin As COrigin, command As String, paramCount As Byte, params() As String)
    m_socket.sendLine "PONG :" & params(0)
End Sub

Public Sub parserError(origin As COrigin, command As String, paramCount As Byte, params() As String)
    m_statusWindowText.addEvent "IRC_ERROR", makeStringArray(params(0))
End Sub

Public Sub parserKill(origin As COrigin, command As String, paramCount As Byte, params() As String)
    m_expectingDisconnect = True
    m_statusWindowText.addEvent "KILLED", makeStringArray(origin.nickname, params(1))
End Sub

Public Sub parserNotice(origin As COrigin, command As String, paramCount As Byte, params() As String)
    m_statusWindowText.addEvent "NOTICE", makeStringArray(origin.nickname, params(1))
End Sub

Public Sub parserWelcome(origin As COrigin, command As String, paramCount As Byte, params() As String)
    m_registered = True

    m_currentNickname = params(0)
    m_serverName = origin.nickname
    m_statusWindow.switchbarTab.caption = m_serverName
    m_client.redrawSwitchbarTab m_statusWindow.switchbarTab

    m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
    m_statusWindowText.addEvent "WELCOME", makeStringArray(origin.nickname, params(0))
    
    registered
End Sub

Public Sub parserNicknameInUse(origin As COrigin, command As String, paramCount As Byte, params() As String)
    If Not m_registered Then
        If m_triedBackupNickname Then
            m_statusWindowText.addEvent "NICKNAME_IN_USE_PREREG2", makeStringArray(params(1))
            m_statusWindowText.inputText = "/nick "
            Beep
        Else
            m_statusWindowText.addEvent "NICKNAME_IN_USE_PREREG", makeStringArray(params(1))
            m_socket.sendLine "NICK " & m_backupNickname
            m_triedBackupNickname = True
        End If
    Else
        m_statusWindowText.addEvent "NICKNAME_IN_USE", makeStringArray(params(1))
    End If
End Sub

Public Sub parserISupport(origin As COrigin, command As String, paramCount As Byte, params() As String)
    parseISupport params(1)
    
    If m_prefixModes.count = 0 Then
        parsePrefixModes getISupportValue("PREFIX")
    End If
    
    If LenB(m_networkName) = 0 Then
        m_networkName = getISupportValue("NETWORK")
        
        If LenB(m_networkName) <> 0 Then
            m_statusWindow.switchbarTab.caption = m_networkName
            m_client.redrawSwitchbarTab m_statusWindow.switchbarTab
        End If
    End If
    
    If m_channelModes.count = 0 Then
        parseChannelModes getISupportValue("CHANMODES")
    End If
    
    If m_maxModes = 0 Then
        m_maxModes = CByte(Val(getISupportValue("MODES")))
    End If
End Sub

Public Sub parserJoin(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    If StrComp(origin.nickname, m_currentNickname, vbTextCompare) = 0 Then
        Set channel = findChannel(params(0))
        
        If channel Is Nothing Then
            Set channel = createChannel(params(0))
            m_client.ShowWindow channel.window
            channel.meJoined
        Else
            channel.reJoined
        End If
    Else
        Set channel = findChannel(params(0))
        
        If Not channel Is Nothing Then
            channel.userJoined origin
        End If
    End If
End Sub

Public Sub parserPart(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel

    Set channel = findChannel(params(0))
    
    If channel Is Nothing Then
        Exit Sub
    End If
    
    If StrComp(origin.nickname, m_currentNickname, vbTextCompare) = 0 Then
        If channel.hopping Then
            channel.reJoining
            m_socket.sendLine "JOIN " & channel.name
        Else
            destroyChannel channel
        End If
    Else
        If paramCount > 1 Then
            channel.userParted origin, params(1)
        Else
            channel.userParted origin
        End If
    End If
End Sub

Public Sub parserKick(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(0))
    
    If channel Is Nothing Then
       Exit Sub
    End If
    
    If StrComp(params(1), m_currentNickname, vbTextCompare) = 0 Then
        channel.meKicked origin, params(2)
    Else
        channel.userKicked origin, params(1), params(2)
    End If
End Sub

Public Sub parserQuit(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    For Each channel In m_channels
        If channel.isIn(origin.nickname) Then
            channel.userQuit origin, params(0)
        End If
    Next channel
End Sub

Public Sub parserNick(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    If StrComp(origin.nickname, m_currentNickname, vbTextCompare) = 0 Then
        m_currentNickname = params(0)
    End If
    
    For Each channel In m_channels
        If channel.isIn(origin.nickname) Then
            channel.userNickChange origin, params(0)
        End If
    Next channel
End Sub

Public Sub parserNames(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(2))
    
    If Not channel Is Nothing Then
        channel.processNamesList params(3)
    End If
End Sub

Public Sub parserEndOfNames(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfNames
    End If
End Sub

Public Sub parserMode(origin As COrigin, command As String, paramCount As Byte, params() As String)
    If paramCount < 2 Then
        Exit Sub
    End If

    Dim channel As CChannel
    
    Set channel = findChannel(params(0))

    If Not channel Is Nothing Then
        If paramCount > 2 Then
            channel.modeChange origin, params(1), params(2)
        Else
            channel.modeChange origin, params(1), vbNullString
        End If
    End If
End Sub

Public Sub parserChannelModeIs(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        If paramCount > 3 Then
            channel.modeIs params(2), params(3)
        Else
            channel.modeIs params(2)
        End If
    End If
End Sub

Public Sub parserTopicIs(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.topicIs params(2)
    End If
End Sub

Public Sub parserTopicWhoTime(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.topicWhoTime params(2), Val(params(3))
    End If
End Sub

Public Sub parserTopic(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(0))
    
    If Not channel Is Nothing Then
        channel.topicChange origin, params(1)
    End If
End Sub

Public Sub parserPrivmsg(origin As COrigin, command As String, paramCount As Byte, params() As String)
    processMessage origin, False, paramCount, params
End Sub

Public Sub parserBanList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "b", params(2)
    End If
End Sub

Public Sub parserEndOfBanList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "b"
    End If
End Sub

Public Sub parserExList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "e", params(2)
    End If
End Sub

Public Sub parserEndOfExList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "e"
    End If
End Sub

Public Sub parserInvexList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "I", params(2)
    End If
End Sub

Public Sub parserEndOfInvexList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "I"
    End If
End Sub

Public Sub parserAList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "a", params(2)
    End If
End Sub

Public Sub parserEndOfAList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "a"
    End If
End Sub

Public Sub parserQList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "q", params(2)
    End If
End Sub

Public Sub parserEndOfQList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "q"
    End If
End Sub

Private Sub processMessage(origin As COrigin, notice As Boolean, paramCount As Byte, params() As String)
    Dim channel As CChannel
    Dim query As CQuery
    Dim prefix As String
    Dim channelName As String
    
    channelName = params(0)
    
    If getPrefixIndex(left$(channelName, 1)) <> 0 Then
        prefix = left$(channelName, 1)
        channelName = Mid$(channelName, 2)
    End If

    Set channel = findChannel(channelName)
    
    If channel Is Nothing Then
        Set query = findQuery(params(0))
    End If
    
    Dim text As String
    
    text = params(1)
    
    If left$(text, 1) = Chr$(1) Then
        Dim ctcpCommand As String
        Dim ctcpParam As String
        
        extractCtcp text, ctcpCommand, ctcpParam
        
        If StrComp(ctcpCommand, "ACTION", vbTextCompare) = 0 Then
            If Not channel Is Nothing Then
                channel.userEmote origin, ctcpParam, prefix
            Else
                If query Is Nothing Then
                    Set query = createQuery(origin.nickname)
                    m_client.ShowWindow query.window
                End If
                
                query.userEmote origin, ctcpParam
            End If
        Else
            If Not channel Is Nothing Then
                channel.userCtcp origin, ctcpCommand, ctcpParam, prefix
            Else
            End If
        End If
    Else
        If Not channel Is Nothing Then
            channel.userMessage origin, text, False, prefix
        Else
            If query Is Nothing Then
                Set query = createQuery(origin.nickname)
                m_client.ShowWindow query.window
            End If
        
            query.userMessage origin, text
        End If
    End If
End Sub

Private Sub extractCtcp(text As String, ByRef command As String, ByRef param As String)
    Dim inParam As Boolean
    
    Do While LenB(text) <> 0
        If left$(text, 1) = Chr$(1) Then
            If LenB(command) <> 0 Then
                Exit Sub
            End If
        ElseIf left$(text, 1) = " " Then
            If LenB(param) = 0 Then
                inParam = True
            Else
                param = param & left$(text, 1)
            End If
        Else
            If inParam Then
                param = param & left$(text, 1)
            Else
                command = command & left$(text, 1)
            End If
        End If
        
        text = Mid$(text, 2)
    Loop
End Sub

Private Function createChannel(name As String) As CChannel
    Dim channelWindow As ctlWindowChannel
    Dim channel As New CChannel
    
    Set channelWindow = m_client.createNewWindow("swiftirc.ctlWindowChannel", "channel")
    
    channel.init name, channelWindow, Me
    channelWindow.init Me, channel
    
    m_channels.Add channel, LCase$(name)
    m_client.addChannelTab Me, channel
    
    Set createChannel = channel
End Function

Private Function createQuery(name As String) As CQuery
    Dim queryWindow As ctlWindowQuery
    Dim query As New CQuery
    
    Set queryWindow = m_client.createNewWindow("swiftIrc.ctlWindowQuery", "query")
    query.init name, queryWindow, Me
    queryWindow.init Me, query
    
    m_queries.Add query, LCase$(name)
    m_client.addQueryTab Me, query
    
    Set createQuery = query
End Function

Private Sub destroyChannel(channel As CChannel)
    m_client.removeTab channel.window.switchbarTab
    m_client.destroyWindow channel.window
    m_channels.Remove LCase$(channel.name)
End Sub

Private Sub destroyQuery(query As CQuery)
    m_client.removeTab query.window.switchbarTab
    m_client.destroyWindow query.window
    m_queries.Remove LCase$(query.name)
End Sub

Public Sub partChannel(channel As CChannel)
    m_socket.sendLine "PART " & channel.name
    destroyChannel channel
End Sub

Public Sub closeQuery(query As CQuery)
    destroyQuery query
End Sub

Private Function findChannel(name As String) As CChannel
    On Error Resume Next
    Set findChannel = m_channels.item(LCase$(name))
End Function

Private Function findQuery(name As String) As CQuery
    On Error Resume Next
    Set findQuery = m_queries.item(LCase$(name))
End Function

Private Sub addISupport(name As String, value As String)
    Dim iSupport As New CISupportEntry

    iSupport.init name, value
    
    On Error Resume Next
    m_iSupport.Add iSupport, UCase(name)
End Sub

Private Function getISupportValue(name As String) As String
    On Error Resume Next
    
    getISupportValue = m_iSupport.item(UCase(name)).value
End Function

Private Sub parseISupport(iSupport As String)
    Dim count As Integer
    Dim char As String
    
    Dim name As String
    Dim value As String
    Dim inValue As Boolean
    
    If right$(iSupport, 1) <> " " Then
        iSupport = iSupport & " "
    End If
    
    For count = 1 To Len(iSupport)
        char = Mid$(iSupport, count, 1)
        
        If char = " " Then
            If LenB(name) <> 0 Then
                If LenB(value) <> 0 Then
                    addISupport name, value
                    value = vbNullString
                Else
                    addISupport name, "1"
                End If
                
                name = vbNullString
                inValue = False
                
                If count < Len(iSupport) Then
                    If Mid$(iSupport, count + 1, 1) = ":" Then
                        Exit For
                    End If
                End If
            End If
        ElseIf char = "=" And Not inValue Then
            inValue = True
        Else
            If inValue Then
                value = value & char
            Else
                name = name & char
            End If
        End If
    Next count
End Sub

Private Sub addPrefixMode(mode As String, symbol As String)
    Dim prefixMode As New CPrefixMode
    
    prefixMode.init mode, symbol
    m_prefixModes.Add prefixMode
End Sub

Public Function comparePrefix(prefix1 As String, prefix2 As String) As Integer
    Dim index1 As Byte
    Dim index2 As Byte
    
    index1 = getPrefixIndex(prefix1)
    index2 = getPrefixIndex(prefix2)
    
    If index1 = 0 And index2 = 0 Then
        comparePrefix = 0
    ElseIf index1 = 0 Then
        comparePrefix = 1
    ElseIf index2 = 0 Then
        comparePrefix = -1
    ElseIf index1 = index2 Then
        comparePrefix = 0
    ElseIf index1 < index2 Then
        comparePrefix = -1
    Else
        comparePrefix = 1
    End If
End Function

Private Function getPrefixIndex(symbol As String) As Byte
    Dim count As Integer
    
    For count = 1 To m_prefixModes.count
        If m_prefixModes.item(count).symbol = symbol Then
            getPrefixIndex = count
            Exit Function
        End If
    Next count
End Function

Private Function getPrefixModeIndex(mode As String) As Byte
    Dim count As Integer
    
    For count = 1 To m_prefixModes.count
        If m_prefixModes.item(count).mode = mode Then
            getPrefixModeIndex = count
            Exit Function
        End If
    Next count
End Function

Public Function isPrefixMode(mode As String) As Boolean
    isPrefixMode = (getPrefixModeIndex(mode) <> 0)
End Function

Public Function getPrefixModeBySymbol(symbol As String) As String
    Dim prefix As CPrefixMode
    
    Set prefix = findPrefixBySymbol(symbol)
    
    If Not prefix Is Nothing Then
        getPrefixModeBySymbol = prefix.mode
    End If
End Function

Public Function getEffectivePrefix(modes As String) As String
    Dim count As Integer
    Dim highest As Byte
    Dim index As Byte
    
    For count = 1 To Len(modes)
        index = getPrefixModeIndex(Mid$(modes, count, 1))
        
        If index < highest Or highest = 0 Then
            highest = index
        End If
    Next count
    
    If highest <> 0 Then
        getEffectivePrefix = m_prefixModes.item(highest).symbol
    End If
End Function

Private Function findPrefixBySymbol(symbol As String) As CPrefixMode
    Dim count As Integer
    
    For count = 1 To m_prefixModes.count
        If m_prefixModes.item(count).symbol = symbol Then
            Set findPrefixBySymbol = m_prefixModes.item(count)
            Exit Function
        End If
    Next count
End Function

Private Sub parsePrefixModes(prefixes As String)
    Dim count As Integer
    Dim char As String
    
    Dim modes As String
    Dim symbols As String
    
    Dim inModes As Boolean
    
    For count = 1 To Len(prefixes)
        char = Mid$(prefixes, count, 1)
        
        If char = "(" Then
            inModes = True
        ElseIf char = ")" Then
            inModes = False
        Else
            If inModes Then
                modes = modes & char
            Else
                symbols = symbols & char
            End If
        End If
    Next count
    
    For count = 1 To Len(modes)
        If count <= Len(symbols) Then
            addPrefixMode Mid$(modes, count, 1), Mid$(symbols, count, 1)
        End If
    Next count
End Sub

Private Sub parseChannelModes(modes As String)
    Dim modeType As eChannelModeType
    Dim channelMode As CChannelMode
    Dim count As Long
    Dim mode As String
    
    For count = 1 To Len(modes)
        mode = Mid$(modes, count, 1)
        
        If mode = "," Then
            If modeType <> cmtUnknown Then
                modeType = modeType + 1
            End If
        Else
            Set channelMode = New CChannelMode
            channelMode.init mode, modeType
            m_channelModes.Add channelMode
        End If
    Next count
End Sub

Public Function getListModes() As String
    Dim count As Long
    
    For count = 1 To m_channelModes.count
        If m_channelModes.item(count).modeType = cmtList Then
            getListModes = getListModes & m_channelModes.item(count).mode
        End If
    Next count
End Function

Public Function getChannelModeType(mode As String) As eChannelModeType
    Dim count As Long
    
    For count = 1 To m_channelModes.count
        If m_channelModes.item(count).mode = mode Then
            getChannelModeType = m_channelModes.item(count).modeType
            Exit Function
        End If
    Next count
    
    getChannelModeType = cmtUnknown
End Function

Private Sub m_socket_disconnected()
    disconnected
End Sub

Private Sub m_socket_error(number As Integer, description As String)
    If m_connected Then
        Select Case number
            Case mswinsocklib.ErrorConstants.sckConnectionReset
                m_statusWindowText.addEvent "ERROR_DISCONNECT", _
                    makeStringArray(textManager.getText("ERR_CONNECT_RESET"))
            Case mswinsocklib.ErrorConstants.sckNetReset
                m_statusWindowText.addEvent "ERROR_DISCONNECT", _
                    makeStringArray(textManager.getText("ERR_CONNECT_NETRESET"))
            Case mswinsocklib.ErrorConstants.sckNoBufferSpace
                m_statusWindowText.addEvent "ERROR_DISCONNECT", _
                    makeStringArray(textManager.getText("ERR_CONNECT_BUFSPACE"))
            Case mswinsocklib.ErrorConstants.sckOutOfMemory
                m_statusWindowText.addEvent "ERROR_DISCONNECT", _
                    makeStringArray(textManager.getText("ERR_CONNECT_MEMORY"))
            Case Else
                m_statusWindowText.addEvent "ERROR_DISCONNECT", makeStringArray("Code: " & number & " - " & description)
        End Select
        
        disconnected
    Else
        Select Case number
            Case mswinsocklib.ErrorConstants.sckHostNotFound
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_DNS"))
            Case mswinsocklib.ErrorConstants.sckHostNotFoundTryAgain
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_DNS"))
            Case mswinsocklib.ErrorConstants.sckNoData
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_DNS"))
            Case mswinsocklib.ErrorConstants.sckTimedout
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_TIMEOUT"))
            Case mswinsocklib.ErrorConstants.sckConnectionRefused
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_REFUSED"))
            Case mswinsocklib.ErrorConstants.sckConnectAborted
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_ABORTED"))
            Case mswinsocklib.ErrorConstants.sckNoBufferSpace
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_BUFSPACE"))
            Case mswinsocklib.ErrorConstants.sckOutOfMemory
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, textManager.getText("ERR_CONNECT_MEMORY"))
            Case Else
                m_statusWindowText.addEvent "ERROR_CONNECT", _
                    makeStringArray(m_serverHost, "Code: " & number & " - " & description)
        End Select
        
        connectFailed
    End If
End Sub

Private Sub m_statusWindow_textInput(text As String)
    textInput m_statusWindow, text
End Sub

Public Sub textInput(source As ITextWindow, line As String)
    line = LTrim(line)
    
    If left$(line, 1) <> "/" Then
        Exit Sub
    End If
    
    Dim command As String
    
    extractCommand line, command
    
    Dim inputParser As CInputParser
    
    Set inputParser = findInputParser(command)
    
    If inputParser Is Nothing Then
        m_socket.sendLine command & " " & line
        Exit Sub
    End If
    
    Dim params() As String
    Dim paramCount As Byte
    
    paramCount = extractCommandParams(line, inputParser.maxParams, params)
    
    If paramCount >= inputParser.minParams Then
        CallByName Me, inputParser.func, VbMethod, source, command, paramCount, params
    End If
End Sub

Public Sub commandJoin(source As ITextWindow, command As String, paramCount As Byte, params() As String)
    If paramCount > 1 Then
        m_socket.sendLine "JOIN " & params(0) & " " & params(1)
    Else
        m_socket.sendLine "JOIN " & params(0)
    End If
End Sub

Public Sub commandTest(source As ITextWindow, command As String, paramCount As Byte, params() As String)
    source.addText params(0)
End Sub

Public Sub commandOptions(source As ITextWindow, command As String, paramCount As Byte, params() As String)
    openOptions m_client, Me
End Sub

Public Sub commandDebug(source As ITextWindow, command As String, paramCount As Byte, params() As String)
    If m_debugWindow Is Nothing Then
        Set m_debugWindow = m_client.createNewWindow("swiftirc.ctlWindowGenericText", "debug")
        Set m_debugWindowText = m_debugWindow
        
        m_debugWindow.init Me
        m_debugWindow.switchbarTab = m_client.addGenericTab(Me, m_debugWindow, "Debug")
        m_client.ShowWindow m_debugWindow
    End If
End Sub

Public Sub commandServer(source As ITextWindow, command As String, paramCount As Byte, params() As String)
    Dim flags As String
    Dim server As String
    Dim port As Long
    Dim password As String
    
    If paramCount = 0 Then
        processServerCommand source, flags, server, port, password
        Exit Sub
    End If

    If left$(params(0), 1) = "-" Then
        If Len(params(0)) > 1 Then
            flags = Mid$(params(0), 2)
        End If
    Else
        If InStr(params(0), ":") <> 0 Then
            splitServerPort params(0), server, port
        Else
            server = params(0)
        End If
    End If
    
    If paramCount > 1 Then
        If LenB(server) <> 0 Then
            If port <> 0 Then
                password = params(1)
            Else
                port = Val(params(1))
            End If
        Else
            If InStr(params(1), ":") <> 0 Then
                splitServerPort params(1), server, port
            Else
                server = params(1)
            End If
        End If
    Else
        processServerCommand source, flags, server, port, password
        Exit Sub
    End If
    
    If paramCount > 2 Then
        If port <> 0 Then
            password = params(2)
        Else
            port = Val(params(2))
        End If
    Else
        processServerCommand source, flags, server, port, password
        Exit Sub
    End If
    
    If paramCount > 3 Then
        If LenB(password) <> 0 Then
            password = params(3)
        End If
    End If
    
    processServerCommand source, flags, server, port, password
End Sub

Private Sub processServerCommand(source As ITextWindow, flags As String, server As String, _
    port As Long, password As String)
    
    Dim profile As CServerProfile
    
    If InStr(flags, "m") <> 0 Then
        Dim newSession As CSession
        
        Set newSession = m_client.newSession
        
        If LenB(server) <> 0 Then
            Set profile = serverProfiles.findProfile(server)
            
            If Not profile Is Nothing Then
                newSession.serverProfile = profile
            Else
                newSession.serverHost = server
                newSession.primaryNickname = m_primaryNickname
                newSession.backupNickname = m_backupNickname
                newSession.realName = m_realName
            End If
            
            If port <> 0 Then
                newSession.serverPort = port
            End If
            
            If LenB(password) <> 0 Then
                newSession.serverPassword = password
            End If
            
            m_client.ShowWindow newSession.statusWindow
            newSession.connect
        Else
            newSession.serverProfile = m_serverProfile
            m_client.ShowWindow newSession.statusWindow
            newSession.connect
        End If
    Else
        If LenB(server) <> 0 Then
            Set profile = serverProfiles.findProfile(server)
            
            If Not profile Is Nothing Then
                Me.serverProfile = profile
            Else
                Me.serverHost = server
            End If
            
            If port <> 0 Then
                Me.serverPort = port
            End If
           
            If LenB(password) <> 0 Then
                Me.serverPassword = password
            End If
        End If
        
        m_client.ShowWindow m_statusWindow
        connect
    End If
End Sub

Private Sub splitServerPort(text As String, ByRef server As String, ByRef port As Long)
    Dim params() As String
    
    params = Split(text, ":", 2)
    
    If UBound(params) = 0 Then
        server = params(0)
        port = 0
        Exit Sub
    End If
    
    server = params(0)
    port = Val(params(1))
End Sub

Private Sub extractCommand(ByRef line As String, ByRef command As String)
    Dim slashes As Integer
    
    Do While LenB(line) <> 0
        Select Case left$(line, 1)
            Case "/"
                slashes = slashes + 1
            Case " "
                If LenB(command) <> 0 Then
                    line = Mid$(line, 2)
                    Exit Sub
                End If
            Case Else
                command = command & left$(line, 1)
        End Select
        
        line = Mid$(line, 2)
    Loop
End Sub

Private Function extractCommandParams(ByRef line As String, maxParams As Byte, ByRef params() As String) As Byte
    Dim temp As String
    Dim paramCount As Byte

    Do While LenB(line) <> 0
        If left$(line, 1) = " " Then
            If LenB(temp) <> 0 Then
                paramCount = paramCount + 1
                ReDim Preserve params(paramCount - 1)
                params(paramCount - 1) = temp
                temp = vbNullString
            End If
        Else
            temp = temp & left$(line, 1)
        End If
        
        If paramCount + 1 >= maxParams Then
            Do While (left$(line, 1) = " ")
                line = Mid$(line, 2)
            Loop
            
            If LenB(line) = 0 Then
                extractCommandParams = paramCount
                Exit Function
            End If
            
            paramCount = paramCount + 1
            ReDim Preserve params(paramCount - 1)
            params(paramCount - 1) = line
            
            extractCommandParams = paramCount
            Exit Function
        End If
        
        line = Mid$(line, 2)
    Loop
    
    If LenB(temp) <> 0 Then
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount - 1)
        params(paramCount - 1) = temp
    End If
    
    extractCommandParams = paramCount
End Function

Private Sub initInputParsers()
    addInputParser "JOIN", 1, 2, "commandJoin"
    addInputParser "TEST", 0, 1, "commandTest"
    addInputParser "OPTIONS", 0, 0, "commandOptions"
    addInputParser "DEBUG", 0, 0, "commandDebug"
    addInputParser "SERVER", 0, 4, "commandServer"
End Sub

Private Sub addInputParser(command As String, minParams As Byte, maxParams As Byte, func As String)
    Dim inputParser As New CInputParser
    
    inputParser.init command, minParams, maxParams, func
    
    m_inputParsers.Add inputParser, LCase$(command)
End Sub

Private Function findInputParser(command As String) As CInputParser
    On Error Resume Next
    Set findInputParser = m_inputParsers.item(LCase$(command))
End Function

Public Sub sendLine(line As String)
    m_socket.sendLine line
End Sub

Public Sub init()
    Set m_reconnectTimer = m_client.getTimer
    m_reconnectTimer.Enabled = False
    m_reconnectTimer.Interval = RECONNECT_DELAY
End Sub

Private Sub Class_Initialize()
    initParsers
    initInputParsers
    Set m_socket = New CSocket
End Sub

Public Sub deInit()
    Dim channel As CChannel
    Dim query As CQuery
    
    For Each channel In m_channels
        destroyChannel channel
    Next channel
    
    For Each query In m_queries
        destroyQuery query
    Next query
    
    m_client.releaseTimer m_reconnectTimer
    
    m_socket.closeSocket
End Sub

Private Sub m_socket_connected()
    m_connected = True
    Call login
End Sub

Private Sub m_socket_readable()
    Dim buffer As String
    
    m_socket.receive buffer
    parse buffer
End Sub

Public Sub channelTabbing(source As swiftIRc.ctlTextInput, text As String, start As Long, length As Long)
    Dim matchLength As Long
    Dim count As Long
    Dim channel As CChannel
    
    If m_lastChannelTabIndex <> 0 Then
        If m_lastChannelTab = text Then
            matchLength = Len(m_lastChannelTabMatch)
        
            For count = m_lastChannelTabIndex + 1 To m_channels.count
                Set channel = m_channels.item(count)
            
                If LCase(left$(channel.name, matchLength)) = m_lastChannelTabMatch Then
                    m_lastChannelTab = channel.name
                    m_lastChannelTabIndex = count
                    
                    source.replaceText start, length, channel.name
                    Exit Sub
                End If
            Next count
            
            If m_lastChannelTabIndex > 1 Then
                For count = 1 To m_lastChannelTabIndex - 1
                    Set channel = m_channels.item(count)
            
                    If LCase(left$(channel.name, matchLength)) = m_lastChannelTabMatch Then
                        m_lastChannelTab = channel.name
                        m_lastChannelTabIndex = count
                        
                        source.replaceText start, length, channel.name
                        Exit Sub
                    End If
                Next count
            End If
            
            Exit Sub
        End If
    End If
    
    m_lastChannelTabMatch = LCase$(text)
    matchLength = Len(text)
    
    For count = 1 To m_channels.count
        Set channel = m_channels.item(count)
            
        If LCase(left$(channel.name, matchLength)) = m_lastChannelTabMatch Then
            m_lastChannelTab = channel.name
            m_lastChannelTabIndex = count
            
            source.replaceText start, length, channel.name
            Exit For
        End If
    Next count
End Sub

Public Sub registered()
    Dim channels As String
    Dim keys As String

    If Not m_serverProfile Is Nothing Then
        If m_serverProfile.enableAutoJoin Then
            Dim ajChannel As CAutoJoinChannel
            Dim count As Long
            
            For count = 1 To m_serverProfile.autoJoinChannelCount
                Set ajChannel = m_serverProfile.autoJoinChannel(count)
            
                If findChannel(ajChannel.channel) Is Nothing Then
                    If LenB(ajChannel.key) <> 0 Then
                        keys = keys & ajChannel.key & ","
                    End If
                    
                    channels = channels & ajChannel.channel
                End If
            Next count
        End If
    End If
    
    Dim channel As CChannel
    
    For Each channel In m_channels
        If LenB(channel.key) <> 0 Then
            keys = keys & channel.key & ","
        End If
        
        channels = channels & channel.name & ","
    Next channel
    
    m_socket.sendLine "JOIN " & channels & " " & keys
End Sub

Public Sub sendModeChange(target As String, modes As String, paramString As String)
    Dim params() As String
    Dim paramIndex As Integer
    Dim modeCount As Integer
    Dim count As Long
    Dim buffer As String
    Dim param As String
    Dim modeOn As Boolean
    
    Dim modeBuffer As String
    Dim paramBuffer As String
    
    Dim modeType As eChannelModeType
    
    params = Split(Trim(paramString), " ")
    
    buffer = "MODE " & target & " "
    
    modeOn = True
    
    For count = 1 To Len(modes)
        param = vbNullString
    
        If modeCount >= m_maxModes Then
            sendLine buffer & modeBuffer & " " & paramBuffer
            
            If modeOn = True Then
                modeBuffer = "+"
            Else
                modeBuffer = "-"
            End If
            
            paramBuffer = ""
            modeCount = 0
        End If
        
        Select Case Mid$(modes, count, 1)
            Case "+"
                modeOn = True
            Case "-"
                modeOn = False
            Case Else
                modeCount = modeCount + 1
                
                If getPrefixModeIndex(Mid$(modes, count, 1)) <> 0 Then
                    If paramIndex <= UBound(params) Then
                        param = params(paramIndex)
                    End If
                Else
                    modeType = getChannelModeType(Mid$(modes, count, 1))
                    
                    If modeType = cmtList Then
                        If paramIndex <= UBound(params) Then
                            param = params(paramIndex)
                        End If
                    ElseIf modeType = cmtParam Then
                        If paramIndex <= UBound(params) Then
                            param = params(paramIndex)
                        End If
                    ElseIf modeType = cmtSetOnly Then
                        If modeOn Then
                            If paramIndex <= UBound(params) Then
                                param = params(paramIndex)
                            End If
                        End If
                    End If
                    
                    If LenB(param) <> 0 Then
                        paramIndex = paramIndex + 1
                    End If
                End If
        End Select
        
        If Len(buffer) + (Len(modeBuffer) + 1) + Len(paramBuffer) + Len(param) + 2 > 500 Then
            sendLine buffer & modeBuffer & " " & paramBuffer
            
            modeBuffer = IIf(modeOn = True, "+", "-") & Mid$(modes, count, 1)
            paramBuffer = param
            modeCount = 1
        Else
            modeBuffer = modeBuffer & Mid$(modes, count, 1)
            
            If LenB(param) <> 0 Then
                paramBuffer = paramBuffer & " " & param
            End If
        End If
    Next count
    
    sendLine buffer & modeBuffer & " " & paramBuffer
End Sub
