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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const RECONNECT_DELAY As Long = 10000
Private Const CHANNEL_JOIN_DELAY As Long = 1000

Private m_serverProfile As CServerProfile
Private m_client As swiftIrc.SwiftIrcClient

Private WithEvents m_statusWindow As ctlWindowStatus
Attribute m_statusWindow.VB_VarHelpID = -1
Private m_statusWindowText As ITextWindow
Private m_debugWindow As ctlWindowGenericText
Private m_debugWindowText As ITextWindow

Private WithEvents m_channelListWindow As ctlChannelList
Attribute m_channelListWindow.VB_VarHelpID = -1
Private m_listSynced As Boolean

Private m_whoisText As ITextWindow

Private WithEvents m_reconnectTimer As VB.timer
Attribute m_reconnectTimer.VB_VarHelpID = -1
Private WithEvents m_channelJoinTimer As VB.timer
Attribute m_channelJoinTimer.VB_VarHelpID = -1

Private WithEvents m_socket As CSwiftSocket
Attribute m_socket.VB_VarHelpID = -1
Private m_resolveIdConnect As Long
Private m_addrList As cArrayList
Private m_currentAddrIndex As Long

Private m_parsers As New Collection
Private m_inputParsers As New Collection
Private m_ial As New Collection

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
Private m_channelTypes As String

Private m_lastCtcp As Long

Public Event connected()
Public Event registered()
Public Event nickInUse()
Public Event badNick()

Private m_baseLogPath As String

Private m_channelsToDestroy As New cArrayList

'Channel tabbing/completion
Private m_lastChannelTab As String
Private m_lastChannelTabIndex As Long
Private m_lastChannelTabMatch As String

Private m_waitingOnUserAgreement As Boolean

Private Declare Function mciSendString Lib "winmm.dll" _
  Alias "mciSendStringA" ( _
  ByVal lpszCommand As String, _
  ByVal lpszReturnString As String, _
  ByVal cchReturnLength As Long, _
  ByVal hwndCallback As Long) As Long
  

Public Property Get baseLogPath() As String
    baseLogPath = m_baseLogPath
End Property

Public Property Get serverProfile() As CServerProfile
    Set serverProfile = m_serverProfile
End Property

Public Property Let serverProfile(newValue As CServerProfile)
    If newValue Is Nothing Then
        Exit Property
    End If

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
    
    If LenB(m_serverProfile.name) <> 0 Then
        m_statusWindow.switchbartab.caption = m_serverProfile.name
    Else
        m_statusWindow.switchbartab.caption = m_serverHost
    End If

    m_client.switchbar.redrawTab m_statusWindow.switchbartab
    
    If settings.setting("enableLogging", estBoolean) Then
        If settings.setting("logDirectories", estBoolean) Then
            If settings.setting("logDirectoriesProfile", estBoolean) _
                And Not m_serverProfile Is serverProfiles.quickConnectProfile Then
                
                m_baseLogPath = g_userPath & LOG_DIR & sanitizeFilename(m_serverProfile.name) & "\"
            
                If Dir(m_baseLogPath, vbDirectory) = vbNullString Then
                    MkDir m_baseLogPath
                End If
                
                If settings.setting("logStatus", estBoolean) Then
                    m_statusWindowText.textview.logName = m_baseLogPath & "status"
                    m_statusWindowText.textview.enableLogging = True
                    m_statusWindowText.textview.writeEntireBuffer
                End If
            End If
        End If
    End If
End Property

Public Property Get client() As swiftIrc.SwiftIrcClient
    Set client = m_client
End Property

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
    Set m_client = newValue
End Property

Public Property Get statusWindow() As ctlWindowStatus
    Set statusWindow = m_statusWindow
End Property

Public Property Let statusWindow(newValue As ctlWindowStatus)
    Set m_statusWindow = newValue
    Set m_statusWindowText = newValue
    
    If settings.setting("enableLogging", estBoolean) Then
        If Not settings.setting("logDirectories", estBoolean) Then
            m_statusWindowText.textview.logName = m_baseLogPath & "status"
            m_statusWindowText.textview.enableLogging = True
        End If
    End If
End Property

Public Property Get serverHost() As String
    serverHost = m_serverHost
End Property

Public Property Let serverHost(newValue As String)
    m_serverHost = newValue
    Set m_addrList = Nothing
    m_currentAddrIndex = 0
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

Public Sub agreementAccepted()
    If m_waitingOnUserAgreement Then
        m_waitingOnUserAgreement = False
        connect
    End If
End Sub

Public Sub connect()
    If Not settings.acceptedFirstUse Then
        m_waitingOnUserAgreement = True
        m_client.showUserAgreement
        Exit Sub
    End If

    m_reconnectTimer.enabled = False

    If m_connected Then
        disconnect
    End If
    
    If LenB(m_serverHost) = 0 Then
        Exit Sub
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
    
    If Not m_addrList Is Nothing Then
        m_socket.connect m_addrList.item(m_currentAddrIndex).addr, m_serverPort
    Else
        Dim addr As Long
    
        addr = inet_addr(m_serverHost)
    
        If addr <> INADDR_NONE Then
            m_socket.connect addr, m_serverPort
        Else
            m_resolveIdConnect = m_socket.resolveAddrByName(m_serverHost)
        End If
    End If
End Sub

Public Sub disconnect()
    m_expectingDisconnect = True
    m_socket.sendLine "QUIT"
    m_socket.closeSocket
    disconnected
End Sub

Private Sub login()
    m_triedBackupNickname = False
    
    m_statusWindowText.addEvent "CONNECTED", makeStringArray(m_serverHost)

    If LenB(m_serverPassword) <> 0 Then
        sendLine "PASS :" & m_serverPassword
    End If
    
    sendLine "USER " & m_username & " * 0 :" & m_realName
    sendLine "NICK " & m_currentNickname
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
    
    m_prefixModes.clear
    m_channelModes.clear
    m_maxModes = 0
    
    m_registered = False
    m_connected = False
    
    If Not m_expectingDisconnect Then
        If Not m_serverProfile Is Nothing Then
            If m_serverProfile.enableReconnect Then
                m_statusWindowText.addEvent "RECONNECTING_IN", makeStringArray(RECONNECT_DELAY / _
                    1000)
                m_reconnectTimer.enabled = True
            End If
        End If
    End If
    
    m_expectingDisconnect = False
End Sub

Private Sub connectFailed()
    If Not m_serverProfile Is Nothing Then
        If m_serverProfile.enableConnectRetry Then
            If Not m_addrList Is Nothing Then
                If m_currentAddrIndex < m_addrList.count Then
                    m_currentAddrIndex = m_currentAddrIndex + 1
                Else
                    Set m_addrList = Nothing
                    m_currentAddrIndex = 0
                End If
            End If
            
            m_statusWindowText.addEvent "RECONNECTING_IN", makeStringArray(RECONNECT_DELAY / 1000)
            m_reconnectTimer.enabled = True
        End If
    End If
End Sub

Private Sub m_channelJoinTimer_Timer()
    joinChannels
    m_channelJoinTimer.enabled = False
End Sub

Private Sub m_channelListWindow_joinChannel(name As String)
    sendLine "JOIN " & name
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
    
    Dim parsedOrigin As New COrigin
    parsedOrigin.parseOrigin origin
    
    updateIal parsedOrigin
    
    Dim parser As CParser
    
    Set parser = findParser(command)
    
    If parser Is Nothing Then
        Exit Sub
    End If
    
    Dim params() As String
    Dim paramCount As Byte
    
    paramCount = extractIrcParams(line, parser.maxParams, params)
    
    If paramCount < parser.minParams Then
        Exit Sub
    End If
    
    CallByName Me, parser.func, VbMethod, parsedOrigin, command, paramCount, params
End Sub

Private Sub extractIrcHeader(ByRef line As String, ByRef origin As String, ByRef command As String)
    Dim currentParam As Byte
    
    origin = vbNullString
    command = vbNullString
    
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
        ElseIf LenB(origin) = 0 And Mid$(line, 1, 1) = ":" Then
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

Private Function extractIrcParams(ByRef line As String, maxParams As Byte, ByRef params() As String) As Integer
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
            
            extractIrcParams = paramCount
            
            Exit Function
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
                extractIrcParams = paramCount
                Exit Function
            End If
            
            If left$(line, 1) = ":" Then
                line = Mid$(line, 2)
                extractIrcParams = paramCount
            End If
            
            If LenB(line) = 0 Then
                Exit Function
            End If
            
            paramCount = paramCount + 1
            ReDim Preserve params(paramCount - 1)
            params(paramCount - 1) = RTrim(line)
            extractIrcParams = paramCount
            Exit Function
        End If
        
        line = Mid$(line, 2)
    Loop
    
    If LenB(temp) <> 0 Then
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount - 1)
        params(paramCount - 1) = temp
    End If
    
    extractIrcParams = paramCount
End Function

Private Function addParser(ByVal name As String, ByVal minParams As Byte, ByVal maxParams As Byte, _
    ByVal func As String)
    
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

    addParser "NICK", 1, 1, "parserNick"

    addParser "001", 2, 2, "parserWelcome"
    addParser "005", 2, 2, "parserISupport"
    addParser "NOTICE", 2, 2, "parserNotice"
    addParser "JOIN", 1, 1, "parserJoin"
    addParser "PART", 1, 2, "parserPart"
    addParser "KICK", 3, 3, "parserKick"
    addParser "QUIT", 0, 1, "parserQuit"
    
    addParser "TOPIC", 2, 2, "parserTopic"
    
    addParser RPL_TOPIC, 3, 3, "parserTopicIs"
    addParser RPL_TOPICWHOTIME, 4, 4, "parserTopicWhoTime"
    
    addParser "353", 4, 4, "parserNames"
    
    addParser ERR_NICKNAMEINUSE, 3, 3, "parserNicknameInUse"
    addParser ERR_NOSUCHNICK, 3, 3, "parserNoSuchNick"
    
    addParser "PRIVMSG", 2, 2, "parserPrivmsg"
    addParser "MODE", 2, 3, "parserMode"
    
    addParser RPL_LISTSTART, 3, 3, "parserListStart"
    addParser RPL_LIST, 3, 4, "parserList"
    addParser RPL_LISTEND, 2, 2, "parserListEnd"
    
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
    
    addParser RPL_WHOISUSER, 6, 6, "parserWhoisUser"
    addParser RPL_WHOISSERVER, 4, 4, "parserWhoisServer"
    addParser RPL_WHOISOPERATOR, 3, 3, "parserWhoisOperator"
    addParser RPL_WHOISHELPOP, 3, 3, "parserWhoisHelpop"
    addParser RPL_WHOISSECURE, 3, 3, "parserWhoisSecure"
    addParser RPL_WHOISBOT, 3, 3, "parserWhoisBot"
    addParser RPL_HASFILTER, 3, 3, "parserWhoisFilter"
    addParser RPL_HASFILTEROVER, 3, 3, "parserWhoisFilterOver"
    addParser RPL_WHOISIDLE, 5, 5, "parserWhoisIdle"
    addParser RPL_ENDOFWHOIS, 3, 3, "parserEndOfWhois"
    addParser RPL_WHOISCHANNELS, 3, 3, "parserWhoisChannels"
    addParser RPL_WHOISREGNICK, 3, 3, "parserWhoisRegNick"
    addParser RPL_AWAY, 3, 3, "parserAway"
    
    addParser ERR_BANNEDFROMCHAN, 3, 3, "parserBannedFromChan"
    addParser ERR_NEEDREGGEDNICK, 3, 3, "ParserNeedReggedNick"
    addParser ERR_BADCHANNELKEY, 3, 3, "parserBadChannelKey"
    addParser ERR_INVITEONLYCHAN, 3, 3, "parserInviteOnlyChan"
    addParser ERR_CHANNELISFULL, 3, 3, "parserChannelIsFull"
    addParser ERR_LINKCHANNEL, 2, 2, "parserLinkChannel"
    addParser ERR_TOOMANYJOINS, 3, 3, "parserTooManyJoins"
    addParser ERR_TOOMANYCHANNELS, 3, 3, "parserTooManyChannels"
    addParser ERR_SECUREONLYCHAN, 3, 3, "parserSecureOnlyChan"
    addParser ERR_OPERONLY, 2, 2, "parserOperOnly"
    addParser ERR_ADMONLY, 2, 2, "parserOperOnly"
    
    addParser ERR_UNKNOWNCOMMAND, 3, 3, "parserUnknownCommand"
    addParser ERR_ERRONEUSNICKNAME, 3, 3, "parserErroneusNickname"
    
    ' 10 Nov 2011
    addParser RPL_WHOREPLY, 8, 8, "parserWhoList"
    addParser RPL_ENDOFWHO, 3, 3, "parserEndOfWhoList"
    
    addParser RPL_WHOWASUSER, 6, 6, "parserWhowasHost"
    addParser ERR_WASNOSUCHNICK, 3, 3, "parserWhowasUnknown"
    addParser RPL_ENDOFWHOWAS, 3, 3, "parserEndOfWhowas"
    
    addParser RPL_SILENCELIST, 3, 3, "parserSilenceList"
    addParser RPL_ENDOFSILENCELIST, 2, 2, "parserEndOfSilenceList"
    
    addParser ERR_CHGNICKMODEN, 2, 2, "parserChgNickModeN"
    addParser ERR_CHGNICKMODEB, 3, 3, "parserChgNickModeb"
    
    addParser RPL_HELPOPTITLE, 2, 2, "parserHelpopTitle"
    addParser RPL_HELPOPTEXT, 2, 2, "parserHelpopText"
    
    addParser RPL_NOWAWAY, 2, 2, "parserMarkedAway"
    addParser RPL_UNAWAY, 2, 2, "parserUnmarkedAway"
    
    addParser RPL_INVITELIST, 2, 2, "parserInviteList"
    addParser RPL_ENDOFINVITELIST, 2, 2, "parserEndOfInviteList"
    
    addParser "SILENCE", 1, 1, "parserSilenceModify"
    addParser "INVITE", 2, 2, "parserInvitationReceived"
    
    addParser RPL_INVITING, 3, 3, "parserInviteUser"
End Sub

Public Sub parserPing(origin As COrigin, command As String, paramCount As Byte, params() As String)
    sendLine "PONG :" & params(0)
End Sub

Public Sub parserError(origin As COrigin, command As String, paramCount As Byte, params() As String)
    m_statusWindowText.addEvent "IRC_ERROR", makeStringArray(params(0))
End Sub

Public Sub parserKill(origin As COrigin, command As String, paramCount As Byte, params() As String)
    m_expectingDisconnect = True
    m_statusWindowText.addEvent "KILLED", makeStringArray(origin.nickname, params(1))
End Sub

Public Sub parserNotice(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    processMessage origin, True, paramCount, params
End Sub

Public Sub parserWelcome(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_registered = True

    m_currentNickname = params(0)
    m_serverName = origin.nickname
    m_statusWindow.switchbartab.caption = m_serverName
    m_client.switchbar.redrawTab m_statusWindow.switchbartab

    m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
    m_statusWindowText.addEvent "WELCOME", makeStringArray(origin.nickname, params(0))
    
    registered
End Sub

Public Sub parserNicknameInUse(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    If Not m_registered Then
        If m_triedBackupNickname Then
            m_statusWindowText.addEvent "NICKNAME_IN_USE_PREREG2", makeStringArray(params(1))
            m_statusWindowText.inputText = "/nick "
            Beep
        Else
            m_statusWindowText.addEvent "NICKNAME_IN_USE_PREREG", makeStringArray(params(1))
            sendLine "NICK " & m_backupNickname
            m_triedBackupNickname = True
        End If
    Else
        m_statusWindowText.addEvent "NICKNAME_IN_USE", makeStringArray(params(1))
    End If
End Sub

Public Sub parserNoSuchNick(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim activeText As ITextWindow
    
    Set activeText = m_client.activeTextWindow
    
    If Not activeText Is Nothing Then
        If activeText.session Is Me Then
            activeText.addEvent "NO_SUCH_NICK", makeStringArray(params(1))
        End If
    End If
End Sub

Public Sub parserISupport(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    parseISupport params(1)
    
    If m_prefixModes.count = 0 Then
        parsePrefixModes getISupportValue("PREFIX")
    End If
    
    If LenB(m_networkName) = 0 Then
        m_networkName = getISupportValue("NETWORK")
        
        If LenB(m_networkName) <> 0 Then
            m_statusWindow.switchbartab.caption = m_networkName
            m_client.switchbar.redrawTab m_statusWindow.switchbartab
            
            If settings.setting("enableLogging", estBoolean) Then
                If settings.setting("logDirectories", estBoolean) Then
                    If Not settings.setting("logDirectoriesProfile", estBoolean) Or _
                        m_serverProfile Is Nothing Or m_serverProfile Is serverProfiles.quickConnectProfile Then
                        
                        m_baseLogPath = g_userPath & LOG_DIR & sanitizeFilename(m_networkName) & "\"
                        
                        If Dir(m_baseLogPath, vbDirectory) = vbNullString Then
                            MkDir m_baseLogPath
                        End If
                        
                        If settings.setting("logStatus", estBoolean) Then
                            m_statusWindowText.textview.logName = m_baseLogPath & "status"
                            m_statusWindowText.textview.enableLogging = True
                            m_statusWindowText.textview.writeEntireBuffer
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If m_channelModes.count = 0 Then
        parseChannelModes getISupportValue("CHANMODES")
    End If
    
    If m_maxModes = 0 Then
        m_maxModes = CByte(Val(getISupportValue("MODES")))
    End If
    
    If LenB(m_channelTypes) = 0 Then
        m_channelTypes = getISupportValue("CHANTYPES")
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
            rejoinChannel channel
        Else
            Dim count As Long
            
            For count = 1 To m_channelsToDestroy.count
                If m_channelsToDestroy.item(count) Is channel Then
                    m_channelsToDestroy.Remove count
                    Exit For
                End If
            Next count
            
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
            If paramCount = 0 Then
                channel.userQuit origin
            Else
                channel.userQuit origin, params(0)
            End If
        End If
    Next channel
End Sub

Public Sub parserNick(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    If StrComp(origin.nickname, m_currentNickname, vbTextCompare) = 0 Then
        m_currentNickname = params(0)
        m_statusWindowText.addEvent "ME_NICK_CHANGE", makeStringArray(params(0))
    End If
    
    For Each channel In m_channels
        If channel.isIn(origin.nickname) Then
            channel.userNickChange origin, params(0)
        End If
    Next channel
    
    
    If Not findQuery(params(0)) Is Nothing Then
        Exit Sub
    End If
    
    Dim query As CQuery
    
    Set query = findQuery(origin.nickname)
    
    If Not query Is Nothing Then
        query.nickChange origin, params(0)
        m_queries.Remove LCase$(origin.nickname)
        m_queries.Add query, LCase$(query.name)
    End If
End Sub

Public Sub parserNames(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(2))
    
    If Not channel Is Nothing Then
        channel.processNamesList params(3)
    End If
End Sub

Public Sub parserEndOfNames(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfNames
    End If
End Sub

Public Sub parserWhoisUser(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Set m_whoisText = m_client.activeTextWindow
    
    If Not m_whoisText.session Is Me Then
        Set m_whoisText = m_statusWindow
    End If
    
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_USER", makeStringArray(params(1), params(2), params(3), _
            params(5))
    End If
End Sub

Public Sub parserWhoisChannels(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_CHANNELS", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisServer(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_SERVER", makeStringArray(params(1), params(2), params(3))
    End If
End Sub

Public Sub parserWhoisRegNick(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_REGNICK", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisOperator(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_OPERATOR", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisHelpOp(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_OPERATOR", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisSecure(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisBot(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisFilter(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisFilterOver(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoisIdle(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_IDLE", makeStringArray(params(1), timeString(Val(params(2))), _
            formatTime(Val(params(3))))
    End If
End Sub

Public Sub parserEndOfWhois(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOIS_END", makeStringArray(params(1), params(2))
        Set m_whoisText = Nothing
    End If
End Sub

Public Sub parserAway(origin As COrigin, command As String, paramCount As Byte, params() As String)
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "AWAY", makeStringArray(params(1), params(2))
    Else
        Dim activeText As ITextWindow
        
        Set activeText = m_client.activeTextWindow
        
        If Not activeText.session Is Me Then
            Set activeText = m_statusWindow
        End If
        
        activeText.addEvent "AWAY", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserBannedFromChan(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.IText.addEvent "ERR_BANNEDFROMCHAN", makeStringArray(params(1))
    End If
    
    m_statusWindowText.addEvent "ERR_BANNEDFROMCHAN", makeStringArray(params(1))
End Sub

Public Sub parserNeedReggedNick(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.IText.addEvent "ERR_NEEDREGGEDNICK", makeStringArray(params(1))
    End If
    
    m_statusWindowText.addEvent "ERR_NEEDREGGEDNICK", makeStringArray(params(1))
End Sub

Public Sub parserInviteOnlyChan(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.IText.addEvent "ERR_INVITEONLYCHAN", makeStringArray(params(1))
    End If
    
    m_statusWindowText.addEvent "ERR_INVITEONLYCHAN", makeStringArray(params(1))
End Sub

Public Sub parserBadChannelKey(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.IText.addEvent "ERR_BADCHANNELKEY", makeStringArray(params(1))
    End If
    
    m_statusWindowText.addEvent "ERR_BADCHANNELKEY", makeStringArray(params(1))
End Sub

Public Sub parserChannelIsFull(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.IText.addEvent "ERR_CHANNELISFULL", makeStringArray(params(1))
    End If
    
    m_statusWindowText.addEvent "ERR_CHANNELISFULL", makeStringArray(params(1))
End Sub

Public Sub parserTooManyJoins(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.IText.addEvent "ERR_TOOMANYJOINS", makeStringArray(params(1))
    End If
    
    m_statusWindowText.addEvent "ERR_TOOMANYJOINS", makeStringArray(params(1))
End Sub

Public Sub parserTooManyChannels(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    m_statusWindowText.addEvent "ERR_TOOMANYCHANNELS", makeStringArray(params(1))
End Sub

Public Sub parserSecureOnlyChan(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)

    m_statusWindowText.addEvent "ERR_SECUREONLYCHAN", makeStringArray(params(1))
End Sub

Public Sub parserLinkChannel(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
End Sub

Public Sub parserNoPrivileges(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "ERR_NOPRIVILEGES", makeStringArray(params(1))
End Sub

Public Sub parserOperOnly(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
End Sub

Public Sub parserUnknownCommand(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1) & " " & params(2))
End Sub

Public Sub parserErroneusNickname(origin As COrigin, command As String, paramCount As Byte, _
    params() As String)
    m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1) & " " & params(2))
End Sub

Public Sub parserMode(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(0))

    If Not channel Is Nothing Then
        If paramCount > 2 Then
            channel.modeChange origin, params(1), params(2)
        Else
            channel.modeChange origin, params(1), vbNullString
        End If
    Else
        If StrComp(params(0), m_currentNickname, vbTextCompare) = 0 Then
            m_statusWindowText.addEvent "ME_MODE_CHANGE", makeStringArray(origin.nickname, params(1))
        End If
    End If
End Sub

Public Sub parserChannelModeIs(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
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

Public Sub parserTopicIs(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.topicIs params(2)
    End If
End Sub

Public Sub parserTopicWhoTime(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
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

Public Sub parserPrivmsg(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    
    processMessage origin, False, paramCount, params
End Sub

Public Sub parserBanList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "b", params(2)
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "BAN_LIST", makeStringArray(params(1), params(2), params(3), _
            formatTime(params(4)))
    End If
End Sub

Public Sub parserEndOfBanList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "b"
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "END_OF_BAN_LIST", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserExList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "e", params(2)
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "EX_LIST", makeStringArray(params(1), params(2), params(3), _
            formatTime(params(4)))
    End If
End Sub

Public Sub parserEndOfExList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "e"
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "END_OF_EX_LIST", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserInvexList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "I", params(2)
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "INVEX_LIST", makeStringArray(params(1), params(2), params(3), _
            formatTime(params(4)))
    End If
End Sub

Public Sub parserEndOfInvexList(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "I"
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "END_OF_INVEX_LIST", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserAList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "a", params(2)
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "A_LIST", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserEndOfAList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "a"
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "END_OF_A_LIST", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserQList(origin As COrigin, command As String, paramCount As Byte, params() As String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.modeListEntry "q", params(2)
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "Q_LIST", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserEndOfQList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    Set channel = findChannel(params(1))
    
    If Not channel Is Nothing Then
        channel.endOfModeList "q"
    End If
    
    If Not CCOpen() Then
        m_statusWindowText.addEvent "END_OF_Q_LIST", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserWhoList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Set m_whoisText = m_client.activeTextWindow
    
    If Not m_whoisText.session Is Me Then
        Set m_whoisText = m_statusWindow
    End If
    
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHO_LIST", makeStringArray(params(1), params(5), params(2), _
            params(3), params(6), params(7))
    End If
End Sub

Public Sub parserEndOfWhoList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Set m_whoisText = m_client.activeTextWindow
    
    If Not m_whoisText.session Is Me Then
        Set m_whoisText = m_statusWindow
    End If
    
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHO_END_OF_LIST", makeStringArray(params(2))
    End If
End Sub

Public Sub parserWhowasHost(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Set m_whoisText = m_client.activeTextWindow
    
    If Not m_whoisText.session Is Me Then
        Set m_whoisText = m_statusWindow
    End If
    
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOWAS_HOST", makeStringArray(params(1), params(2), params(3), params(5))
    End If
End Sub

Public Sub parserWhowasUnknown(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Set m_whoisText = m_client.activeTextWindow
    
    If Not m_whoisText.session Is Me Then
        Set m_whoisText = m_statusWindow
    End If
    
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOWAS_UNKNOWN", makeStringArray(params(1), params(2))
    End If
End Sub

Public Sub parserEndOfWhowas(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    Set m_whoisText = m_client.activeTextWindow
    
    If Not m_whoisText.session Is Me Then
        Set m_whoisText = m_statusWindow
    End If
    
    If Not m_whoisText Is Nothing Then
        m_whoisText.addEvent "WHOWAS_END", makeStringArray(params(2))
    End If
End Sub

Public Sub parserSilenceModify(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "SILENCE_MODIFY", makeStringArray(params(0))
End Sub

Public Sub parserInvitationReceived(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    
    If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_INVITE) Then
        Exit Sub
    End If
    
    If m_client.activeTextWindow.session Is Me Then
        m_client.activeTextWindow.addEvent "INVITATION_RECEIVED", makeStringArray(params(1), _
            origin.nickname, origin.username, origin.hostname)
    Else
        m_statusWindowText.addEvent "INVITATION_RECEIVED", makeStringArray(params(1), _
            origin.nickname, origin.username, origin.hostname)
    End If
End Sub

Public Sub parserListStart(rigin As COrigin, command As String, paramCount As Byte, params() As _
    String)

    If m_channelListWindow Is Nothing Then
        showList
    Else
        m_channelListWindow.clear
        m_client.ShowWindow m_channelListWindow
    End If
End Sub

Public Sub parserList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    
    If Not m_channelListWindow Is Nothing Then
        m_channelListWindow.insertChannel params(1), Val(params(2)), params(3)
    End If
End Sub

Public Sub parserListEnd(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_listSynced = True
    m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
End Sub

Public Sub parserSilenceList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "SILENCE_LIST", makeStringArray(params(2))
End Sub

Public Sub parserEndOfSilenceList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "END_OF_SILENCE_LIST", makeStringArray(params(1))
End Sub

Public Sub parserChgNickModeN(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "ERR_CHGNICK_MODEN", makeStringArray(params(1))
End Sub

Public Sub parserChgNickModeB(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "ERR_CHGNICK_MODEB", makeStringArray(params(1), params(2))
End Sub

Public Sub parserHelpopTitle(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "HELPOP_TITLE", makeStringArray(params(1))
End Sub

Public Sub parserHelpopText(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "HELPOP_TEXT", makeStringArray(params(1))
End Sub

Public Sub parserMarkedAway(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "MARKED_AWAY", makeStringArray(params(1))
End Sub

Public Sub parserUnmarkedAway(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "NO_LONGER_MARKED_AWAY", makeStringArray(params(1))
End Sub

Public Sub parserInviteList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "INVITE_LIST", makeStringArray(params(1))
End Sub

Public Sub parserEndOfInviteList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "END_OF_INVITE_LIST", makeStringArray(params(1))
End Sub

Public Sub parserInviteUser(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
    m_statusWindowText.addEvent "INVITE_USER", makeStringArray(params(1), params(2))
End Sub

Private Sub processMessage(origin As COrigin, notice As Boolean, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    Dim query As CQuery
    Dim prefix As String
    Dim channelName As String
    Dim activeText As ITextWindow
    
    On Error GoTo processMessage_Error

    Set activeText = m_client.activeTextWindow
    
    If Not activeText Is Nothing Then
        If Not activeText.session Is Me Then
            Set activeText = m_statusWindow
        End If
    Else
        Set activeText = m_statusWindow
    End If
    
    channelName = params(0)
    
    If getPrefixIndex(left$(channelName, 1)) <> 0 Then
        prefix = left$(channelName, 1)
        channelName = Mid$(channelName, 2)
    End If

    Set channel = findChannel(channelName)
    
    If channel Is Nothing Then
        If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_PRIVATE) Then
            Exit Sub
        End If
    
        Set query = findQuery(origin.nickname)
    Else
        If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_CHANNEL) Then
            Exit Sub
        End If
    End If
    
    Dim text As String
    Dim highlight As Boolean
    
    text = params(1)
    
    If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_CODES) Then
        text = stripFormattingCodes(text)
    End If
    
    If settings.enableHighlighting Then
        If settings.highlightNickname Then
            If InStr(1, text, m_currentNickname, vbTextCompare) Then
                highlight = True
            End If
        End If
        
        If Not highlight Then
            Dim count As Long
            
            For count = 1 To highlights.highlightCount
                If InStr(1, text, highlights.highlightItem(count).text, vbTextCompare) Then
                    highlight = True
                    Exit For
                End If
            Next count
        End If
    End If
    
    If highlight Then
        highlighted
    End If
    
    If left$(text, 1) = Chr$(1) Then
        Dim ctcpCommand As String
        Dim ctcpParam As String
        
        extractCtcp text, ctcpCommand, ctcpParam
        
        If StrComp(ctcpCommand, "ACTION", vbTextCompare) = 0 Then
            If Not channel Is Nothing Then
                channel.userEmote origin, ctcpParam, prefix, highlight
            Else
                If InStr(m_channelTypes, left$(channelName, 1)) <> 0 Then
                    Exit Sub
                End If
            
                If query Is Nothing Then
                    Set query = createQuery(origin.nickname)
                End If
                
                query.userEmote origin, ctcpParam, highlight
            End If
        Else
            If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_CTCP) Then
                Exit Sub
            End If
        
            If notice Then
                processCtcpReply origin, ctcpCommand, ctcpParam
            Else
                processCtcp origin, params(0), ctcpCommand, ctcpParam
            End If
        End If
    Else
        'Handled after ctcp processing to avoid unintentionally
        'blocking ctcp replies (which are really just notices)
        If notice And ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_NOTICE) Then
            Exit Sub
        End If
    
        If Not channel Is Nothing Then
            channel.userMessage origin, text, notice, prefix, highlight
        Else
            If notice Then
                m_statusWindowText.addEvent "NOTICE", makeStringArray(origin.nickname, text)
                m_client.switchbar.tabActivity m_statusWindow.switchbartab, tasMessage
                
                If Not activeText Is Nothing And Not activeText Is m_statusWindow Then
                    activeText.addEvent "NOTICE", makeStringArray(origin.nickname, text)
                    
                    Dim tabWindow As ITabWindow
                    Set tabWindow = activeText
                    
                    m_client.switchbar.tabActivity tabWindow.getTab, tasMessage
                End If
            Else
                If InStr(m_channelTypes, left$(channelName, 1)) <> 0 Then
                    Exit Sub
                End If
            
                If query Is Nothing Then
                    Set query = createQuery(origin.nickname)
                End If
                
                query.userMessage origin, text, highlight
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

processMessage_Error:
    handleError "processMessage", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub highlighted()
    If settings.setting("highlightCustomSound", estBoolean) Then
        PlaySound settings.setting("highlightSoundPath", estString)
    Else
        Beep
    End If
End Sub

Private Sub PlaySound(soundPath As String)
    Dim command As String
    Dim alias As String
    Dim path As String
    Dim ext As String
    Dim audioType As String
    
    Randomize
    alias = "MySound" & Rnd
    path = soundPath
    ext = right$(path, 4)
    
    If ext = ".mp3" Then
        audioType = "mpegvideo"
    ElseIf ext = ".wav" Then
        audioType = "waveaudio"
    Else
        audioType = vbNullString
        Beep
    End If
    
    If audioType <> vbNullString Then
        command = "open """ & path & _
            """ type " & audioType & _
            " alias " & alias
            
        If mciSendString(command, 0, 0, 0) = 0 Then
            command = "play " & alias & _
            " from 0"
            mciSendString command, 0, 0, 0
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

Private Sub processCtcp(origin As COrigin, target As String, command As String, param As String)
    If GetTickCount - m_lastCtcp < 2500 Then
        Exit Sub
    End If
    
    Dim prefix As String
    Dim channelName As String
    Dim channel As CChannel
    
    If getPrefixIndex(left$(target, 1)) <> 0 Then
        prefix = left$(target, 1)
        channelName = Mid$(target, 2)
    End If

    Set channel = findChannel(channelName)
    
    Dim activeWindow As ITextWindow
    
    Set activeWindow = m_client.activeTextWindow
    
    If Not activeWindow Is Nothing Then
        If Not activeWindow.session Is Me Then
            Set activeWindow = m_statusWindow
        End If
    Else
        Set activeWindow = m_statusWindow
    End If
    
    Select Case UCase$(command)
        Case "PING"
            sendCtcpReply origin.nickname, "PING", param
            
            If Not channel Is Nothing Then
                channel.userCtcp origin, command, vbNullString, prefix
            Else
                activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command)
            End If
        Case "VERSION"
            sendCtcpReply origin.nickname, "VERSION", "SwiftIRC client V" & m_client.getVersion
            
            If Not channel Is Nothing Then
                channel.userCtcp origin, command, vbNullString, prefix
            Else
                activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command)
            End If
        Case "TIME"
            sendCtcpReply origin.nickname, "TIME", format(Now, "yyyy-mm-dd hh:MM:ss")
            
            If Not channel Is Nothing Then
                channel.userCtcp origin, "TIME", vbNullString, prefix
            Else
                activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command)
            End If
        Case Else
            If Not channel Is Nothing Then
                channel.userCtcp origin, command, param, prefix
            Else
                activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command, param)
            End If
    End Select
    
    m_lastCtcp = GetTickCount
End Sub

Private Sub processCtcpReply(origin As COrigin, command As String, param As String)
    Dim ctcpParam As String

    Select Case UCase$(command)
        Case "PING"
            Dim since As Long
            
            If LenB(param) <> 0 Then
                If Val(param) <= 2 ^ 31 Then
                    since = GetTickCount - Val(param)
                    ctcpParam = since & "ms"
                Else
                    ctcpParam = param
                End If
            End If
        Case Else
            ctcpParam = param
    End Select

    If Not m_client.activeTextWindow Is Nothing Then
        If m_client.activeTextWindow.session Is Me Then
            m_client.activeTextWindow.addEvent "CTCP_REPLY_RECEIVED", makeStringArray(origin.nickname, command, ctcpParam)
        Else
            m_statusWindowText.addEvent "CTCP_REPLY_RECEIVED", makeStringArray(origin.nickname, command, ctcpParam)
        End If
    Else
        m_statusWindowText.addEvent "CTCP_REPLY_RECEIVED", makeStringArray(origin.nickname, command, ctcpParam)
    End If
End Sub

Public Sub sendCtcp(target As String, ctcp As String)
    Select Case UCase$(ctcp)
        Case "PING"
            sendLine "PRIVMSG " & target & " :" & Chr$(1) & "PING " & GetTickCount & Chr$(1)
        Case Else
            sendLine "PRIVMSG " & target & " :" & Chr$(1) & ctcp & Chr$(1)
    End Select
    
    If Not m_client.activeTextWindow Is Nothing Then
        If m_client.activeTextWindow.session Is Me Then
            m_client.activeTextWindow.addEvent "CMD_CTCP_SENT", makeStringArray(target, ctcp)
        Else
            m_statusWindowText.addEvent "CMD_CTCP_SENT", makeStringArray(target, ctcp)
        End If
    Else
        m_statusWindowText.addEvent "CMD_CTCP_SENT", makeStringArray(target, ctcp)
    End If
End Sub

Private Sub sendCtcpReply(target As String, command As String, param As String)
    sendLine "NOTICE " & target & " :" & Chr$(1) & command & " " & param & Chr$(1)
End Sub

Private Function createChannel(name As String) As CChannel
    Dim channelWindow As ctlWindowChannel
    Dim channel As New CChannel
    
    Set channelWindow = m_client.createNewWindow("swiftirc.ctlWindowChannel", "channel")
    
    channel.init name, channelWindow, Me
    
    channelWindow.init Me, channel
    channelWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, channelWindow, sboChannel, _
        channel.name, g_iconSBChannel)
    
    m_channels.Add channel, LCase$(name)
    Set createChannel = channel
End Function

Public Sub query(name As String, Optional text As String = vbNullString)
    Dim query As CQuery
    
    Set query = findQuery(name)
    
    If Not query Is Nothing Then
        m_client.ShowWindow query.window
    Else
        Set query = createQuery(name)
        m_client.ShowWindow query.window
    End If
    
    If LenB(text) <> 0 Then
        query.textEntered text
    End If
End Sub

Private Function createQuery(name As String) As CQuery
    Dim queryWindow As ctlWindowQuery
    Dim query As New CQuery
    
    Set queryWindow = m_client.createNewWindow("swiftIrc.ctlWindowQuery", "query")
    query.init name, queryWindow, Me
    queryWindow.init Me, query
    queryWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, queryWindow, sboQuery, query.name, g_iconSBQuery)

    m_queries.Add query, LCase$(name)
    Set createQuery = query
End Function

Private Sub destroyChannel(channel As CChannel)
    m_client.removeTab channel.window.switchbartab
    m_client.destroyWindow channel.window
    m_channels.Remove LCase$(channel.name)
    channel.deInit
End Sub

Private Sub destroyQuery(query As CQuery)
    m_client.removeTab query.window.switchbartab
    m_client.destroyWindow query.window
    m_queries.Remove LCase$(query.name)
    query.deInit
End Sub

Public Sub partChannel(channel As CChannel)
    sendLine "PART " & channel.name
    destroyChannel channel
End Sub

Public Sub closeQuery(query As CQuery)
    destroyQuery query
End Sub

Public Sub closeGenericWindow(genericWindow As ctlWindowGenericText)
    m_client.removeTab genericWindow.switchbartab
    m_client.destroyWindow genericWindow
    
    If genericWindow Is m_debugWindow Then
        Set m_debugWindow = Nothing
        Set m_debugWindowText = Nothing
    End If
End Sub

Private Function findChannel(name As String) As CChannel
    On Error Resume Next
    Set findChannel = m_channels.item(LCase$(name))
End Function

Private Function findQuery(name As String) As CQuery
    On Error Resume Next
    Set findQuery = m_queries.item(LCase$(name))
End Function

Friend Sub joinChannel(name As String)
    sendLine "JOIN " & name
End Sub

Friend Sub rejoinChannel(channel As CChannel)
    channel.reJoining

    If LenB(channel.key) <> 0 Then
        sendLine "JOIN " & channel.name & " " & channel.key
    Else
        sendLine "JOIN " & channel.name
    End If
End Sub

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

Friend Function getUserStyle(nickname As String, prefix As String) As CUserStyle
    If StrComp(nickname, m_currentNickname, vbTextCompare) = 0 _
        And settings.setting("colourMyNick", estBoolean) Then
        
        Dim prefixIndex As Long
        prefixIndex = getPrefixIndex(prefix)
        
        If prefixIndex = 0 Then
            Set getUserStyle = styleMe
        ElseIf prefixIndex <= getPrefixIndex("@") Then
            Set getUserStyle = styleMeOp
        ElseIf prefixIndex <= getPrefixIndex("%") Then
            Set getUserStyle = styleMeHalfop
        ElseIf prefixIndex <= getPrefixIndex("+") Then
            Set getUserStyle = styleMeVoice
        End If
    Else
        Set getUserStyle = getPrefixStyle(prefix)
        
        If getUserStyle Is Nothing Then
            Set getUserStyle = styleNormal
        End If
    End If
End Function

Friend Function getPrefixStyle(symbol As String) As CUserStyle
    Dim count As Long
    
    Dim index As Long
    
    index = getPrefixIndex(symbol)
    
    If index = 0 Then
        Exit Function
    End If
    
    For count = 1 To prefixStyles.count
        If index <= getPrefixIndex(prefixStyles.item(count).symbol) Then
            Set getPrefixStyle = prefixStyles.item(count).style
            Exit Function
        End If
    Next count
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

Private Sub m_socket_closed(errorCode As Long)
    disconnected
End Sub

Private Sub m_socket_connectFailed(errorCode As Long)
    Select Case errorCode
        Case 10060
            m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                textManager.getText("ERR_CONNECT_TIMEOUT"))
        Case 10061
            m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                textManager.getText("ERR_CONNECT_REFUSED"))
        Case 10053
            m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                textManager.getText("ERR_CONNECT_ABORTED"))
        Case 10055
            m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                textManager.getText("ERR_CONNECT_BUFSPACE"))
        Case 8
            m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                textManager.getText("ERR_CONNECT_MEMORY"))
        Case Else
            m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, "Code: " _
                & errorCode)
    End Select
    
    connectFailed
End Sub

Private Sub m_socket_resolveCompleteA(requestId As Long, addrList As cArrayList)
    If requestId = m_resolveIdConnect Then
        Set m_addrList = addrList
        m_currentAddrIndex = 1
        m_socket.connect addrList.item(m_currentAddrIndex).addr, m_serverPort
        m_resolveIdConnect = 0
    End If
End Sub

Private Sub m_socket_resolveFailed(requestId As Long, errorCode As Long)
    m_resolveIdConnect = 0
    m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
        textManager.getText("ERR_CONNECT_DNS"))
        
    connectFailed
End Sub

Private Sub m_statusWindow_textInput(text As String)
    textInput m_statusWindow, text
End Sub

Public Sub textInput(Source As ITextWindow, line As String)
    line = LTrim(line)
    
    If left$(line, 1) <> "/" Then
        Exit Sub
    End If
    
    Dim command As String
    
    extractCommand line, command
    
    Dim inputParser As CInputParser
    
    Set inputParser = findInputParser(command)
    
    If inputParser Is Nothing Then
        sendLine command & " " & line
        Exit Sub
    End If
    
    Dim params() As String
    Dim paramCount As Byte
    
    paramCount = extractCommandParams(line, inputParser.maxParams, params)
    
    If paramCount >= inputParser.minParams Then
        CallByName Me, inputParser.func, VbMethod, Source, command, paramCount, params
    Else
        Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
    End If
End Sub

Public Sub commandList(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    
    If paramCount = 0 Then
        sendLine "LIST"
    ElseIf paramCount = 1 Then
        sendLine "LIST *" & params(0) & "*"
    Else
        sendLine "LIST " & params(0) & " " & params(1)
    End If
End Sub

Public Sub commandJoin(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    
    Dim channelName As String
    Dim key As String
    
    channelName = params(0)
    
    If InStr(m_channelTypes, left$(channelName, 1)) = 0 Then
        channelName = "#" & channelName
    End If
    
    If paramCount > 1 Then
        key = params(1)
    End If
    
    If LenB(key) <> 0 Then
        'Key was provided, no need to see
        'if we have a cached one for the channel.
        
        sendLine "JOIN " & channelName & " " & key
    Else
        Dim channel As CChannel
        
        Set channel = findChannel(channelName)
        
        If Not channel Is Nothing Then
            If Not channel.meIsIn Then
                rejoinChannel channel
            End If
        Else
            sendLine "JOIN " & channelName
        End If
    End If
End Sub

Public Sub commandPart(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    If paramCount = 0 Then
        If TypeOf Source Is ctlWindowChannel Then
            Dim windowChannel As ctlWindowChannel
            Set windowChannel = Source
            partChannel windowChannel.channel
        Else
            Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
        End If
    ElseIf paramCount = 1 Then
        sendLine "PART " & params(0)
    Else
        sendLine "PART " & params(0) & " : " & params(1)
    End If
End Sub

Public Sub commandMsg(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    sendLine "PRIVMSG " & params(0) & " :" & params(1)
    Source.addEvent "CMD_PRIVMSG_SENT", makeStringArray(params(0), params(1))
End Sub

Public Sub commandNotice(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    sendLine "NOTICE " & params(0) & " :" & params(1)
    Source.addEvent "CMD_NOTICE_SENT", makeStringArray(params(0), params(1))
End Sub

Public Sub commandQuery(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    If paramCount = 1 Then
        query params(0)
    Else
        query params(0), params(1)
    End If
End Sub

Public Sub commandMe(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    If TypeOf Source Is ctlWindowChannel Then
        Dim windowChannel As ctlWindowChannel
        
        Set windowChannel = Source
        windowChannel.channel.sendEmote params(0)
    ElseIf TypeOf Source Is ctlWindowQuery Then
        Dim windowQuery As ctlWindowQuery
        
        Set windowQuery = Source
        windowQuery.query.sendEmote params(0)
    Else
        Source.addEvent "CMD_INCOMPATIBLE_WINDOW", makeStringArray(command)
    End If
End Sub

Public Sub commandTest(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    Source.addText params(0)
End Sub

Public Sub commandTopic(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    
    Dim channelWindow As ctlWindowChannel
    
    If paramCount > 1 Then
        sendLine "TOPIC " & params(0) & " :" & params(1)
    ElseIf paramCount > 0 Then
        If TypeOf Source Is ctlWindowChannel Then
            Set channelWindow = Source
            sendLine "TOPIC " & channelWindow.channel.name & " :" & params(0)
        Else
            sendLine "TOPIC " & params(0)
        End If
    Else
        If TypeOf Source Is ctlWindowChannel Then
            Set channelWindow = Source
            sendLine "TOPIC " & channelWindow.channel.name
        End If
    End If
End Sub

Public Sub commandOptions(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    openOptionsDialog m_client
End Sub

Public Sub commandDebug(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    If m_debugWindow Is Nothing Then
        Set m_debugWindow = m_client.createNewWindow("swiftirc.ctlWindowGenericText", "debug")
        Set m_debugWindowText = m_debugWindow
        
        m_debugWindow.init Me, "Debug"
        m_debugWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, m_debugWindow, sboGeneric, "Debug", g_iconSBGeneric)
        m_client.ShowWindow m_debugWindow
    End If
End Sub

Public Sub commandEcho(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    Source.addText params(0)
End Sub

Public Sub commandRaw(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    Source.addEvent "CMD_RAW_SENT", makeStringArray(params(0))
    sendLine params(0)
End Sub

Public Sub commandKick(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    Dim channel As CChannel
    
    If TypeOf Source Is ctlWindowChannel And InStr(m_channelTypes, left$(params(0), 1)) = 0 Then
        Dim channelWindow As ctlWindowChannel
        Set channelWindow = Source
        
        If paramCount > 2 Then
            sendLine "KICK " & channelWindow.channel.name & " " & params(0) & " :" & params(1) & " " & params(2)
        ElseIf paramCount > 1 Then
            sendLine "KICK " & channelWindow.channel.name & " " & params(0) & " :" & params(1)
        Else
            sendLine "KICK " & channelWindow.channel.name & " " & params(0) & " :No reason specified"
        End If
    Else
        If paramCount > 1 Then
            If paramCount > 2 Then
                sendLine "KICK " & params(0) & " " & params(1) & " :" & params(2)
            Else
                sendLine "KICK " & params(0) & " " & params(1) & " :No reason specified"
            End If
        Else
            Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
        End If
    End If
End Sub

Public Sub commandSlap(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    Dim channelWindow As ctlWindowChannel
    
    If TypeOf Source Is ctlWindowChannel Then
        Set channelWindow = Source
        channelWindow.channel.sendEmote "slaps " & params(0) & " around a bit with a large trout"
    Else
        Source.addEvent "CMD_INCOMPATIBLE_WINDOW", makeStringArray(command)
    End If
End Sub

Public Sub commandHuggle(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    Dim channelWindow As ctlWindowChannel
    
    If TypeOf Source Is ctlWindowChannel Then
        Set channelWindow = Source
        channelWindow.channel.sendEmote "huggles " & params(0)
    Else
        Source.addEvent "CMD_INCOMPATIBLE_WINDOW", makeStringArray(command)
    End If
End Sub

Public Sub commandHop(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    
    Dim channelWindow As ctlWindowChannel
    
    If TypeOf Source Is ctlWindowChannel Then
        Set channelWindow = Source
        channelWindow.channel.hopping = True
        
        If paramCount < 1 Then
            channelWindow.channel.hop
        Else
            channelWindow.channel.hop params(0)
        End If
    Else
        If paramCount < 1 Then
            Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
        Else
            Dim channel As CChannel
            
            Set channel = findChannel(params(0))
        
            If Not channel Is Nothing Then
                If paramCount < 2 Then
                    channelWindow.channel.hop
                Else
                    channelWindow.channel.hop params(1)
                End If
            End If
        End If
    End If

End Sub

Private Sub showList()
    Set m_channelListWindow = m_client.createNewWindow("swiftirc.ctlChannelList", "list")
    m_channelListWindow.session = Me
    m_channelListWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, m_channelListWindow, sboChannelList, "Channels", g_iconSBList)
    m_client.ShowWindow m_channelListWindow
End Sub

Friend Sub closeChannelList()
    If Not m_channelListWindow Is Nothing Then
        m_client.removeTab m_channelListWindow.switchbartab
        m_client.destroyWindow m_channelListWindow
        Set m_channelListWindow = Nothing
    End If
End Sub

Public Sub commandClear(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    
    Source.clear
End Sub

Public Sub commandClearAll(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    
    Dim textWindow As ITextWindow
    Dim channel As CChannel
    Dim query As CQuery
    
    m_statusWindowText.clear
    
    For Each channel In m_channels
        Set textWindow = channel.window
        textWindow.clear
    Next channel
    
    For Each query In m_queries
        Set textWindow = query.window
        textWindow.clear
    Next query
    
    If Not m_debugWindowText Is Nothing Then
        m_debugWindowText.clear
    End If
End Sub

Public Sub commandCtcp(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    sendCtcp params(0), params(1)
End Sub

Public Sub commandClose(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    If TypeOf Source Is ctlWindowChannel Then
        Dim channelWindow As ctlWindowChannel
        
        Set channelWindow = Source
        partChannel channelWindow.channel
    ElseIf TypeOf Source Is ctlWindowQuery Then
        Dim queryWindow As ctlWindowQuery
    
        Set queryWindow = Source
        closeQuery queryWindow.query
    ElseIf TypeOf Source Is ctlWindowGenericText Then
        closeGenericWindow Source
    End If
End Sub

Public Sub commandCloseAll(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    
    If TypeOf Source Is ctlWindowChannel Then
        Dim channel As CChannel
        
        For Each channel In m_channels
            partChannel channel
        Next channel
    ElseIf TypeOf Source Is ctlWindowQuery Then
        Dim query As CQuery
        
        For Each query In m_queries
            closeQuery query
        Next query
    End If
End Sub


Public Sub commandReconnect(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    connect
End Sub

Public Sub commandQuit(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    
    m_expectingDisconnect = True
    
    If paramCount > 0 Then
        sendLine "QUIT :" & params(0)
    Else
        sendLine "QUIT"
    End If
End Sub

Public Sub commandIgnore(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
    
    Dim commandFlags As String
    Dim ignoreFlags As String
    Dim mask As String
    
    If paramCount = 0 Then
        m_client.showIgnoreList Me
        Exit Sub
    ElseIf paramCount > 0 Then
        If left$(params(0), 1) = "-" Then
            commandFlags = Mid$(params(0), 2)
            
            If paramCount > 1 Then
                mask = params(1)
                
                If paramCount > 2 Then
                    ignoreFlags = params(2)
                End If
            End If
        Else
            mask = params(0)
            
            If paramCount > 1 Then
                ignoreFlags = params(1)
            End If
        End If
    End If
    
    processIgnoreCommand Source, commandFlags, mask, ignoreFlags
End Sub

Private Sub processIgnoreCommand(Source As ITextWindow, commandFlags As String, mask As String, ignoreFlags As String)
    Dim ignore As CIgnoreItem
    Dim removed As Boolean

    commandFlags = LCase$(commandFlags)

    If commandFlags = "a" Or commandFlags = "add" Or LenB(commandFlags) = 0 Then
        If LenB(mask) <> 0 Then
            Set ignore = New CIgnoreItem
            
            If swiftMatch("*!*@*", mask) Then
                ignore.mask = mask
            Else
                ignore.mask = getIal(mask, ialHost)
            End If
            
            If LenB(ignoreFlags) <> 0 Then
                If Not ignore.parseFlagChars(ignoreFlags) Then
                    Source.addEvent "IGNORE_INVALID_FLAGS", makeStringArray(ignoreFlags)
                    Exit Sub
                End If
            Else
                ignore.flags = IGNORE_ALL
            End If
            
            If swiftMatch("*!*@*", mask) Then
                removed = ignoreManager.removeIgnoreByMask(mask)
            Else
                removed = ignoreManager.removeIgnoreByMask(getIal(mask, ialAll))
            End If
            
            ignoreManager.addIgnore ignore
            
            If removed Then
                Source.addEvent "IGNORE_UPDATED", makeStringArray(ignore.mask, ignore.flagChars)
            Else
                Source.addEvent "IGNORE_ADDED", makeStringArray(ignore.mask, ignore.flagChars)
            End If
            
            ignoreManager.saveIgnoreList
        End If
    ElseIf commandFlags = "l" Or commandFlags = "list" Then
        listIgnores Source
    ElseIf commandFlags = "r" Or commandFlags = "remove" Then
        If LenB(mask) <> 0 Then
            If swiftMatch("*!*@*", mask) Then
                Set ignore = ignoreManager.getIgnoreByMask(mask)
                removed = ignoreManager.removeIgnoreByMask(mask)
            Else
                Set ignore = ignoreManager.getIgnoreByMask(getIal(mask, ialAll))
                removed = ignoreManager.removeIgnoreByMask(getIal(mask, ialAll))
            End If
            
            If removed Then
                Source.addEvent "IGNORE_REMOVED", makeStringArray(ignore.mask)
                ignoreManager.saveIgnoreList
            Else
                Source.addEvent "IGNORE_REMOVE_NOTFOUND", makeStringArray(mask)
            End If
        End If
    ElseIf commandFlags = "c" Or commandFlags = "clear" Then
        ignoreManager.clearIgnores
        Source.addEvent "IGNORE_LIST_CLEARED", makeStringArray("")
        ignoreManager.saveIgnoreList
    Else
        Source.addEvent "IGNORE_INVALID_COMMAND", makeStringArray(commandFlags)
    End If
End Sub

Private Sub listIgnores(window As ITextWindow)
    Dim count As Long

    window.addEvent "IGNORE_LIST_START", makeStringArray("")
    
    For count = 1 To ignoreManager.ignoreCount
        window.addEvent "IGNORE_LIST_ENTRY", _
            makeStringArray(ignoreManager.ignore(count).mask, ignoreManager.ignore(count).flagChars)
    Next count
    
    window.addEvent "IGNORE_LIST_END", makeStringArray("")
End Sub


Public Sub commandServer(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
    Dim flags As String
    Dim server As String
    Dim port As Long
    Dim password As String
    
    If paramCount = 0 Then
        processServerCommand Source, flags, server, port, password
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
        processServerCommand Source, flags, server, port, password
        Exit Sub
    End If
    
    If paramCount > 2 Then
        If port <> 0 Then
            password = params(2)
        Else
            port = Val(params(2))
        End If
    Else
        processServerCommand Source, flags, server, port, password
        Exit Sub
    End If
    
    If paramCount > 3 Then
        If LenB(password) <> 0 Then
            password = params(3)
        End If
    End If
    
    processServerCommand Source, flags, server, port, password
End Sub

Private Sub processServerCommand(Source As ITextWindow, flags As String, server As String, port As _
    Long, password As String)
    
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
                newSession.serverPort = m_serverPort
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
            If Not m_serverProfile Is Nothing Then
                newSession.serverProfile = m_serverProfile
            Else
                newSession.serverHost = m_serverHost
                newSession.serverPort = m_serverPort
                newSession.primaryNickname = m_primaryNickname
                newSession.backupNickname = m_backupNickname
                newSession.realName = m_realName
            End If
            
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

Private Function extractCommandParams(ByRef line As String, maxParams As Byte, ByRef params() As _
    String) As Byte
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
    addInputParser "J", 1, 2, "commandJoin"
    
    addInputParser "TOPIC", 0, 2, "commandTopic"
    
    addInputParser "KICK", 1, 3, "commandKick"
    addInputParser "K", 1, 3, "commandKick"
    
    addInputParser "HOP", 0, 2, "commandHop"
    addInputParser "REJOIN", 0, 2, "commandHop"
    
    addInputParser "ECHO", 1, 1, "commandEcho"
    
    addInputParser "QUOTE", 1, 1, "commandRaw"
    addInputParser "RAW", 1, 1, "commandRaw"
    
    addInputParser "PART", 0, 2, "commandPart"
    
    addInputParser "OPTIONS", 0, 0, "commandOptions"
    addInputParser "DEBUG", 0, 0, "commandDebug"
    addInputParser "SERVER", 0, 4, "commandServer"
    
    addInputParser "RECONNECT", 0, 0, "commandReconnect"
    
    addInputParser "MSG", 2, 2, "commandMsg"
    addInputParser "NOTICE", 2, 2, "commandNotice"
    addInputParser "ME", 1, 1, "commandMe"
    addInputParser "EMOTE", 1, 1, "commandMe"
    addInputParser "ACTION", 1, 1, "commandMe"
    addInputParser "QUERY", 1, 2, "commandQuery"
    
    addInputParser "SLAP", 1, 1, "commandSlap"
    addInputParser "HUGGLE", 1, 1, "commandHuggle"
    
    addInputParser "CLEAR", 0, 0, "commandClear"
    addInputParser "CLEARALL", 0, 0, "commandClearAll"
    
    addInputParser "CTCP", 2, 3, "commandCtcp"
    
    addInputParser "CLOSE", 0, 0, "commandClose"
    addInputParser "CLOSEALL", 0, 0, "commandCloseAll"
    
    addInputParser "LIST", 0, 2, "commandList"
    addInputParser "QUIT", 0, 1, "commandQuit"
    
    addInputParser "UNICODE", 0, 0, "commandUnicode"
    
    addInputParser "IGNORE", 0, 3, "commandIgnore"
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
    If Not m_debugWindowText Is Nothing Then
        m_debugWindowText.addText "-> " & line
    End If
    
    m_socket.sendLine line
End Sub

Public Sub init()
    Set m_reconnectTimer = m_client.getTimer
    Set m_channelJoinTimer = m_client.getTimer
    
    m_reconnectTimer.enabled = False
    m_reconnectTimer.Interval = RECONNECT_DELAY
    
    m_channelJoinTimer.enabled = False
    m_channelJoinTimer.Interval = CHANNEL_JOIN_DELAY
    
    Set m_socket = New CSwiftSocket
    
    m_socket.init m_statusWindow.hwnd
End Sub

Private Sub Class_Initialize()
    initParsers
    initInputParsers
    
    m_baseLogPath = g_userPath & LOG_DIR
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
    m_client.releaseTimer m_channelJoinTimer
    
    m_socket.closeSocket
    m_statusWindow.deInit
    
    If Not m_debugWindow Is Nothing Then
        closeGenericWindow m_debugWindow
    End If
    
    closeChannelList
    
    Set m_statusWindow = Nothing
    Set m_statusWindowText = Nothing
    Set m_whoisText = Nothing
    Set m_client = Nothing
End Sub

Private Sub m_socket_connected()
    m_connected = True
    Call login
End Sub

Private Sub m_socket_readable()
    Dim buf As String
    Dim result As Long
    
    result = m_socket.sockRead(buf)
    
    If result > 0 Then
        parse buf
    End If
End Sub

Public Sub channelTabbing(Source As swiftIrc.ctlTextInput, text As String, start As Long, Length As _
    Long)
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
                    
                    Source.replaceText start, Length, channel.name
                    Exit Sub
                End If
            Next count
            
            If m_lastChannelTabIndex > 1 Then
                For count = 1 To m_lastChannelTabIndex - 1
                    Set channel = m_channels.item(count)
            
                    If LCase(left$(channel.name, matchLength)) = m_lastChannelTabMatch Then
                        m_lastChannelTab = channel.name
                        m_lastChannelTabIndex = count
                        
                        Source.replaceText start, Length, channel.name
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
            
            Source.replaceText start, Length, channel.name
            Exit For
        End If
    Next count
End Sub

Public Sub registered()
    If settings.enableFiltering Then
        sendLine "MODE " & m_currentNickname & " +f"
    End If
    
    Dim channels As String
    
    If Not m_serverProfile Is Nothing Then
        If m_serverProfile.enableAutoIdentify Then
            sendLine "NICKSERV IDENTIFY " & m_serverProfile.nicknamePassword
        End If
    
        If m_serverProfile.enablePerform Then
            runCommands m_serverProfile.perform
        End If
    End If
    
    m_channelJoinTimer.enabled = True
End Sub

Private Sub runCommands(commands As String)
    Dim count As Long
    Dim line As String
    
    For count = 1 To Len(commands)
        If Mid$(commands, count, 1) = vbLf Or Mid$(commands, count, 1) = vbCr Then
            If count < Len(commands) Then
                If Mid$(commands, count + 1, 1) = vbLf Or Mid$(commands, count + 1, 1) = vbCr Then
                    count = count + 1
                End If
            End If
            
            If left$(LTrim(line), 1) <> "/" Then
                line = "/" & line
            End If
            
            line = Replace$(line, "$me", m_currentNickname, 1, -1, vbTextCompare)
            textInput m_statusWindow, line
            line = vbNullString
        Else
            line = line & Mid$(commands, count, 1)
        End If
    Next count
    
    If LenB(line) <> 0 Then
        If left$(LTrim(line), 1) <> "/" Then
            line = "/" & line
        End If
            
        line = Replace$(line, "$me", m_currentNickname, 1, -1, vbTextCompare)
        textInput m_statusWindow, line
    End If
End Sub

Private Sub joinChannels()
    Dim channels As String

    If Not m_serverProfile Is Nothing Then
        If m_serverProfile.enableAutoJoin Then
            Dim ajChannel As CAutoJoinChannel
            Dim count As Long
            
            For count = 1 To m_serverProfile.autoJoinChannelCount
                Set ajChannel = m_serverProfile.autoJoinChannel(count)
            
                If findChannel(ajChannel.channel) Is Nothing Then
                    If Len(channels) + Len(ajChannel.channel) + 8 _
                        >= 510 Then
                        sendLine "JOIN " & channels
                        channels = vbNullString
                    End If
                
                    If LenB(ajChannel.key) <> 0 Then
                        sendLine "JOIN " & channels
                        sendLine "JOIN " & ajChannel.channel & " " & ajChannel.key
                        channels = vbNullString
                    Else
                        channels = channels & ajChannel.channel & ","
                    End If
                End If
            Next count
        End If
    End If
    
    Dim channel As CChannel
    
    For Each channel In m_channels
        If Len(channels) + Len(channel.name) + 8 >= 510 Then
            sendLine "JOIN " & channels
            channels = vbNullString
        End If
        
        If LenB(channel.key) <> 0 Then
            sendLine "JOIN " & channels
            sendLine "JOIN " & channel.name & " " & channel.key
            channels = vbNullString
        Else
            channels = channels & channel.name & ","
        End If
    Next channel
    
    If LenB(channels) <> 0 Then
        sendLine "JOIN " & channels
    End If
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
                End If
                
                If LenB(param) <> 0 Then
                    paramIndex = paramIndex + 1
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

Private Sub updateIal(origin As COrigin)
    Dim ialEntry As CIalEntry
    
    Set ialEntry = findIal(origin.nickname)
    
    If Not ialEntry Is Nothing Then
        ialEntry.nick = origin.nickname
        ialEntry.host = origin.hostname
        ialEntry.ident = origin.username
        Exit Sub
    End If
    
    Set ialEntry = New CIalEntry
    
    ialEntry.nick = origin.nickname
    ialEntry.ident = origin.username
    ialEntry.host = origin.hostname
    
    m_ial.Add ialEntry, LCase$(ialEntry.nick)
End Sub

Private Function findIal(nickname As String) As CIalEntry
    On Error Resume Next
    Set findIal = m_ial.item(LCase$(nickname))
End Function

Public Function getIal(nickname As String, ialType As eIalType) As String
    Dim ialEntry As CIalEntry
    
    Set ialEntry = findIal(nickname)
    
    If ialEntry Is Nothing Then
        getIal = nickname & "!*@*"
        Exit Function
    End If
    
    Select Case ialType
        Case ialAll
            getIal = ialEntry.nick & "!" & ialEntry.ident & "@" & ialEntry.host
        Case ialHost
            getIal = "*!*@" & ialEntry.host
        Case ialIdent
            getIal = "*!*" & ialEntry.ident & "@*"
        Case ialIdentHost
            getIal = "*!*" & ialEntry.ident & "@" & ialEntry.host
        Case Else
            getIal = ialEntry.nick & "!" & ialEntry.ident & "@" & ialEntry.host
    End Select
End Function

Public Function CCOpen() As Boolean
    ' This function will return true if the Channel Central
    ' is opened, false if not
    
    Dim frm As Form
    Dim result As Boolean
    result = False
    
    For Each frm In Forms
        If frm.name = "frmChannelCentral" Then
            result = True
        End If
    Next frm
    
    CCOpen = result
End Function

