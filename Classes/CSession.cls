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
10        baseLogPath = m_baseLogPath
End Property

Public Property Get serverProfile() As CServerProfile
10        Set serverProfile = m_serverProfile
End Property

Public Property Let serverProfile(newValue As CServerProfile)
10        If newValue Is Nothing Then
20            Exit Property
30        End If

40        Set m_serverProfile = newValue
          
50        m_primaryNickname = m_serverProfile.primaryNickname
60        m_backupNickname = m_serverProfile.backupNickname
70        m_username = m_serverProfile.primaryNickname
          
80        m_currentNickname = m_primaryNickname
          
90        If LenB(m_serverProfile.realName) <> 0 Then
100           m_realName = m_serverProfile.realName
110       Else
120           m_realName = m_serverProfile.primaryNickname
130       End If
          
140       m_serverHost = m_serverProfile.hostname
          
150       If m_serverProfile.port <> 0 Then
160           m_serverPort = m_serverProfile.port
170       Else
180           m_serverPort = 6667
190       End If
          
200       m_serverPassword = m_serverProfile.serverPassword
          
210       If LenB(m_serverProfile.name) <> 0 Then
220           m_statusWindow.switchbartab.caption = m_serverProfile.name
230       Else
240           m_statusWindow.switchbartab.caption = m_serverHost
250       End If

260       m_client.switchbar.redrawTab m_statusWindow.switchbartab
          
270       If settings.setting("enableLogging", estBoolean) Then
280           If settings.setting("logDirectories", estBoolean) Then
290               If settings.setting("logDirectoriesProfile", estBoolean) _
                      And Not m_serverProfile Is serverProfiles.quickConnectProfile Then
                      
300                   m_baseLogPath = g_userPath & LOG_DIR & sanitizeFilename(m_serverProfile.name) & "\"
                  
310                   If Dir(m_baseLogPath, vbDirectory) = vbNullString Then
320                       MkDir m_baseLogPath
330                   End If
                      
340                   If settings.setting("logStatus", estBoolean) Then
350                       m_statusWindowText.textview.logName = m_baseLogPath & "status"
360                       m_statusWindowText.textview.enableLogging = True
370                       m_statusWindowText.textview.writeEntireBuffer
380                   End If
390               End If
400           End If
410       End If
End Property

Public Property Get client() As swiftIrc.SwiftIrcClient
10        Set client = m_client
End Property

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
10        Set m_client = newValue
End Property

Public Property Get statusWindow() As ctlWindowStatus
10        Set statusWindow = m_statusWindow
End Property

Public Property Let statusWindow(newValue As ctlWindowStatus)
10        Set m_statusWindow = newValue
20        Set m_statusWindowText = newValue
          
30        If settings.setting("enableLogging", estBoolean) Then
40            If Not settings.setting("logDirectories", estBoolean) Then
50                m_statusWindowText.textview.logName = m_baseLogPath & "status"
60                m_statusWindowText.textview.enableLogging = True
70            End If
80        End If
End Property

Public Property Get serverHost() As String
10        serverHost = m_serverHost
End Property

Public Property Let serverHost(newValue As String)
10        m_serverHost = newValue
20        Set m_addrList = Nothing
30        m_currentAddrIndex = 0
End Property

Public Property Get serverPort() As Long
10        serverPort = m_serverPort
End Property

Public Property Let serverPort(newValue As Long)
10        m_serverPort = newValue
End Property

Public Property Get serverPassword() As String
10        serverPassword = m_serverPassword
End Property

Public Property Let serverPassword(newValue As String)
10        m_serverPassword = newValue
End Property

Public Property Get currentNickname() As String
10        currentNickname = m_currentNickname
End Property

Public Property Get primaryNickname() As String
10        primaryNickname = m_primaryNickname
End Property

Public Property Let primaryNickname(newValue As String)
10        m_primaryNickname = newValue
20        m_currentNickname = newValue
30        m_username = newValue
End Property

Public Property Get backupNickname() As String
10        backupNickname = m_backupNickname
End Property

Public Property Let backupNickname(newValue As String)
10        m_backupNickname = newValue
End Property

Public Property Get username() As String
10        username = m_username
End Property

Public Property Let username(newValue As String)
10        m_username = newValue
End Property

Public Property Get realName() As String
10        realName = m_realName
End Property

Public Property Let realName(newValue As String)
10        m_realName = newValue
End Property

Public Property Get networkName() As String
10        networkName = m_networkName
End Property

Public Property Get connected() As Boolean
10        connected = m_connected
End Property

Public Sub agreementAccepted()
10        If m_waitingOnUserAgreement Then
20            m_waitingOnUserAgreement = False
30            connect
40        End If
End Sub

Public Sub connect()
10        If Not settings.acceptedFirstUse Then
20            m_waitingOnUserAgreement = True
30            m_client.showUserAgreement
40            Exit Sub
50        End If

60        m_reconnectTimer.enabled = False

70        If m_connected Then
80            disconnect
90        End If
          
100       If LenB(m_serverHost) = 0 Then
110           Exit Sub
120       End If

130       m_statusWindowText.addEvent "CONNECTING", makeStringArray(m_serverHost, m_serverPort)
          
140       If m_serverPort = 0 Then
150           m_serverPort = 6667
160       End If
          
170       If LenB(m_username) = 0 Then
180           m_username = m_primaryNickname
190       End If
          
200       If LenB(m_realName) = 0 Then
210           m_realName = m_primaryNickname
220       End If
          
230       If Not m_addrList Is Nothing Then
240           m_socket.connect m_addrList.item(m_currentAddrIndex).addr, m_serverPort
250       Else
              Dim addr As Long
          
260           addr = inet_addr(m_serverHost)
          
270           If addr <> INADDR_NONE Then
280               m_socket.connect addr, m_serverPort
290           Else
300               m_resolveIdConnect = m_socket.resolveAddrByName(m_serverHost)
310           End If
320       End If
End Sub

Public Sub disconnect()
10        m_expectingDisconnect = True
20        m_socket.sendLine "QUIT"
30        m_socket.closeSocket
40        disconnected
End Sub

Private Sub login()
10        m_triedBackupNickname = False
          
20        m_statusWindowText.addEvent "CONNECTED", makeStringArray(m_serverHost)

30        If LenB(m_serverPassword) <> 0 Then
40            sendLine "PASS :" & m_serverPassword
50        End If
          
60        sendLine "USER " & m_username & " * 0 :" & m_realName
70        sendLine "NICK " & m_currentNickname
End Sub

Private Sub disconnected()
10        m_statusWindowText.addEvent "DISCONNECTED", makeStringArray("")
          
          Dim channel As CChannel
          Dim textWnd As ITextWindow
          
20        For Each channel In m_channels
30            channel.disconnected
40        Next channel
          
50        Do While m_iSupport.count <> 0
60            m_iSupport.Remove 1
70        Loop
          
80        m_networkName = vbNullString
          
90        m_prefixModes.clear
100       m_channelModes.clear
110       m_maxModes = 0
          
120       m_registered = False
130       m_connected = False
          
140       If Not m_expectingDisconnect Then
150           If Not m_serverProfile Is Nothing Then
160               If m_serverProfile.enableReconnect Then
170                   m_statusWindowText.addEvent "RECONNECTING_IN", makeStringArray(RECONNECT_DELAY / _
                          1000)
180                   m_reconnectTimer.enabled = True
190               End If
200           End If
210       End If
          
220       m_expectingDisconnect = False
End Sub

Private Sub connectFailed()
10        If Not m_serverProfile Is Nothing Then
20            If m_serverProfile.enableConnectRetry Then
30                If Not m_addrList Is Nothing Then
40                    If m_currentAddrIndex < m_addrList.count Then
50                        m_currentAddrIndex = m_currentAddrIndex + 1
60                    Else
70                        Set m_addrList = Nothing
80                        m_currentAddrIndex = 0
90                    End If
100               End If
                  
110               m_statusWindowText.addEvent "RECONNECTING_IN", makeStringArray(RECONNECT_DELAY / 1000)
120               m_reconnectTimer.enabled = True
130           End If
140       End If
End Sub

Private Sub Class_Terminate()
10        debugLog "CSession terminating"
End Sub

Private Sub m_channelJoinTimer_Timer()
10        joinChannels
20        m_channelJoinTimer.enabled = False
End Sub

Private Sub m_channelListWindow_joinChannel(name As String)
10        sendLine "JOIN " & name
End Sub

Private Sub m_reconnectTimer_Timer()
10        connect
End Sub

Private Sub parse(buffer As String)
          Dim count As Integer
          Dim char As Integer
          Dim last As Integer
          
10        For count = 1 To Len(buffer)
20            char = AscW(Mid$(buffer, count, 1))
              
30            If char = 10 Or char = 13 Then
40                If count < Len(buffer) Then
50                    char = AscW(Mid$(buffer, count + 1, 1))
                      
60                    If char = 10 Or char = 13 Then
70                        count = count + 1
80                    End If
90                End If
                  
100               parseLine
110               m_currentLine = vbNullString
120           Else
130               m_currentLine = m_currentLine & ChrW$(char)
140           End If
150       Next count
End Sub

Private Sub parseLine()
          Dim origin As String
          Dim command As String
          
          Dim line As String
          Dim fullLine As String
          
10        line = m_currentLine
20        fullLine = m_currentLine
30        m_currentLine = vbNullString
          
40        extractIrcHeader line, origin, command
          
50        If LenB(command) = 0 Then
60            Exit Sub
70        End If
          
80        If Not m_debugWindowText Is Nothing Then
90            m_debugWindowText.addText "<- " & fullLine
100       End If
          
          Dim parsedOrigin As New COrigin
110       parsedOrigin.parseOrigin origin
          
120       updateIal parsedOrigin
          
          Dim parser As CParser
          
130       Set parser = findParser(command)
          
140       If parser Is Nothing Then
150           Exit Sub
160       End If
          
          Dim params() As String
          Dim paramCount As Byte
          
170       paramCount = extractIrcParams(line, parser.maxParams, params)
          
180       If paramCount < parser.minParams Then
190           Exit Sub
200       End If
          
210       CallByName Me, parser.func, VbMethod, parsedOrigin, command, paramCount, params
End Sub

Private Sub extractIrcHeader(ByRef line As String, ByRef origin As String, ByRef command As String)
          Dim currentParam As Byte
          
10        origin = vbNullString
20        command = vbNullString
          
30        currentParam = 1
          
40        Do While LenB(line) <> 0
50            If Mid$(line, 1, 1) = " " Then
60                If LenB(command) <> 0 Then
70                    line = Mid$(line, 2)
                      
80                    If LenB(origin) = 0 Then
90                        origin = m_serverName
100                   End If
                      
110                   Exit Sub
120               Else
130                   currentParam = 1
140               End If
150           ElseIf LenB(origin) = 0 And Mid$(line, 1, 1) = ":" Then
160               If LenB(command) = 0 Then
170                   currentParam = 0
180               End If
190           Else
200               If currentParam = 0 Then
210                   origin = origin & Mid$(line, 1, 1)
220               Else
230                   command = command & Mid$(line, 1, 1)
240               End If
250           End If
              
260           line = Mid$(line, 2)
270       Loop
End Sub

Private Function extractIrcParams(ByRef line As String, maxParams As Byte, ByRef params() As String) As Integer
          Dim temp As String
          Dim paramCount As Byte
          
10        Do While LenB(line) <> 0
20            If left$(line, 2) = " :" Then
30                If LenB(temp) <> 0 Then
40                    paramCount = paramCount + 1
50                    ReDim Preserve params(paramCount - 1)
60                    params(paramCount - 1) = temp
70                End If
              
80                line = RTrim(Mid$(line, 3))
90                paramCount = paramCount + 1
100               ReDim Preserve params(paramCount - 1)
110               params(paramCount - 1) = line
                  
120               extractIrcParams = paramCount
                  
130               Exit Function
140           ElseIf left$(line, 1) = " " Then
150               If LenB(temp) <> 0 Then
160                   paramCount = paramCount + 1
170                   ReDim Preserve params(paramCount - 1)
180                   params(paramCount - 1) = temp
190                   temp = vbNullString
200               End If
210           Else
220               temp = temp & left$(line, 1)
230           End If
              
240           If paramCount + 1 >= maxParams Then
250               Do While left$(line, 1) = " "
260                   line = Mid$(line, 2)
270               Loop
                  
280               If LenB(line) = 0 Then
290                   extractIrcParams = paramCount
300                   Exit Function
310               End If
                  
320               If left$(line, 1) = ":" Then
330                   line = Mid$(line, 2)
340                   extractIrcParams = paramCount
350               End If
                  
360               If LenB(line) = 0 Then
370                   Exit Function
380               End If
                  
390               paramCount = paramCount + 1
400               ReDim Preserve params(paramCount - 1)
410               params(paramCount - 1) = RTrim(line)
420               extractIrcParams = paramCount
430               Exit Function
440           End If
              
450           line = Mid$(line, 2)
460       Loop
          
470       If LenB(temp) <> 0 Then
480           paramCount = paramCount + 1
490           ReDim Preserve params(paramCount - 1)
500           params(paramCount - 1) = temp
510       End If
          
520       extractIrcParams = paramCount
End Function

Private Function addParser(ByVal name As String, ByVal minParams As Byte, ByVal maxParams As Byte, _
    ByVal func As String)
          
          Dim parser As CParser
          
10        Set parser = New CParser
          
20        parser.name = name
30        parser.minParams = minParams
40        parser.maxParams = maxParams
50        parser.func = func
          
60        m_parsers.Add parser, LCase$(name)
End Function

Private Function findParser(name As String) As CParser
10        On Error Resume Next
20        Set findParser = m_parsers.item(LCase$(name))
End Function
    
Private Sub initParsers()
10        addParser "PING", 1, 1, "parserPing"
20        addParser "ERROR", 1, 1, "parserError"

30        addParser "KILL", 2, 2, "parserKill"

40        addParser "NICK", 1, 1, "parserNick"

50        addParser "001", 2, 2, "parserWelcome"
60        addParser "005", 2, 2, "parserISupport"
70        addParser "NOTICE", 2, 2, "parserNotice"
80        addParser "JOIN", 1, 1, "parserJoin"
90        addParser "PART", 1, 2, "parserPart"
100       addParser "KICK", 3, 3, "parserKick"
110       addParser "QUIT", 0, 1, "parserQuit"
          
120       addParser "TOPIC", 2, 2, "parserTopic"
          
130       addParser RPL_TOPIC, 3, 3, "parserTopicIs"
140       addParser RPL_TOPICWHOTIME, 4, 4, "parserTopicWhoTime"
          
150       addParser "353", 4, 4, "parserNames"
          
160       addParser ERR_NICKNAMEINUSE, 3, 3, "parserNicknameInUse"
170       addParser ERR_NOSUCHNICK, 3, 3, "parserNoSuchNick"
          
180       addParser "PRIVMSG", 2, 2, "parserPrivmsg"
190       addParser "MODE", 2, 3, "parserMode"
          
200       addParser RPL_LISTSTART, 3, 3, "parserListStart"
210       addParser RPL_LIST, 3, 4, "parserList"
220       addParser RPL_LISTEND, 2, 2, "parserListEnd"
          
230       addParser RPL_CHANNELMODEIS, 3, 4, "parserChannelModeIs"
          
240       addParser RPL_BANLIST, 3, 5, "parserBanList"
250       addParser RPL_ENDOFBANLIST, 3, 3, "parserEndOfBanList"
          
260       addParser RPL_EXLIST, 3, 5, "parserExList"
270       addParser RPL_ENDOFEXLIST, 3, 3, "parserEndOfExList"
          
280       addParser RPL_INVEXLIST, 3, 5, "parserInvexList"
290       addParser RPL_ENDOFINVEXLIST, 3, 3, "parserEndOfInvexList"
          
300       addParser RPL_ALIST, 3, 3, "parserAlist"
310       addParser RPL_ENDOFALIST, 3, 3, "parserEndOfAlist"
          
320       addParser RPL_QLIST, 3, 3, "parserQList"
330       addParser RPL_ENDOFQLIST, 3, 3, "parserEndOfQList"
          
340       addParser RPL_WHOISUSER, 6, 6, "parserWhoisUser"
350       addParser RPL_WHOISSERVER, 4, 4, "parserWhoisServer"
360       addParser RPL_WHOISOPERATOR, 3, 3, "parserWhoisOperator"
370       addParser RPL_WHOISHELPOP, 3, 3, "parserWhoisHelpop"
380       addParser RPL_WHOISSECURE, 3, 3, "parserWhoisSecure"
390       addParser RPL_WHOISBOT, 3, 3, "parserWhoisBot"
400       addParser RPL_HASFILTER, 3, 3, "parserWhoisFilter"
410       addParser RPL_HASFILTEROVER, 3, 3, "parserWhoisFilterOver"
420       addParser RPL_WHOISIDLE, 5, 5, "parserWhoisIdle"
430       addParser RPL_ENDOFWHOIS, 3, 3, "parserEndOfWhois"
440       addParser RPL_WHOISCHANNELS, 3, 3, "parserWhoisChannels"
450       addParser RPL_WHOISREGNICK, 3, 3, "parserWhoisRegNick"
460       addParser RPL_AWAY, 3, 3, "parserAway"
          
470       addParser ERR_BANNEDFROMCHAN, 3, 3, "parserBannedFromChan"
480       addParser ERR_NEEDREGGEDNICK, 3, 3, "ParserNeedReggedNick"
490       addParser ERR_BADCHANNELKEY, 3, 3, "parserBadChannelKey"
500       addParser ERR_INVITEONLYCHAN, 3, 3, "parserInviteOnlyChan"
510       addParser ERR_CHANNELISFULL, 3, 3, "parserChannelIsFull"
520       addParser ERR_LINKCHANNEL, 2, 2, "parserLinkChannel"
530       addParser ERR_TOOMANYJOINS, 3, 3, "parserTooManyJoins"
540       addParser ERR_TOOMANYCHANNELS, 3, 3, "parserTooManyChannels"
550       addParser ERR_SECUREONLYCHAN, 3, 3, "parserSecureOnlyChan"
560       addParser ERR_OPERONLY, 2, 2, "parserOperOnly"
570       addParser ERR_ADMONLY, 2, 2, "parserOperOnly"
          
580       addParser ERR_UNKNOWNCOMMAND, 3, 3, "parserUnknownCommand"
590       addParser ERR_ERRONEUSNICKNAME, 3, 3, "parserErroneusNickname"
          
          ' 10 Nov 2011
600       addParser RPL_WHOREPLY, 8, 8, "parserWhoList"
610       addParser RPL_ENDOFWHO, 3, 3, "parserEndOfWhoList"
          
620       addParser RPL_WHOWASUSER, 6, 6, "parserWhowasHost"
630       addParser ERR_WASNOSUCHNICK, 3, 3, "parserWhowasUnknown"
640       addParser RPL_ENDOFWHOWAS, 3, 3, "parserEndOfWhowas"
          
650       addParser RPL_SILENCELIST, 3, 3, "parserSilenceList"
660       addParser RPL_ENDOFSILENCELIST, 2, 2, "parserEndOfSilenceList"
          
670       addParser ERR_CHGNICKMODEN, 2, 2, "parserChgNickModeN"
680       addParser ERR_CHGNICKMODEB, 3, 3, "parserChgNickModeb"
          
690       addParser RPL_HELPOPTITLE, 2, 2, "parserHelpopTitle"
700       addParser RPL_HELPOPTEXT, 2, 2, "parserHelpopText"
          
710       addParser RPL_NOWAWAY, 2, 2, "parserMarkedAway"
720       addParser RPL_UNAWAY, 2, 2, "parserUnmarkedAway"
          
730       addParser RPL_INVITELIST, 2, 2, "parserInviteList"
740       addParser RPL_ENDOFINVITELIST, 2, 2, "parserEndOfInviteList"
          
750       addParser "SILENCE", 1, 1, "parserSilenceModify"
760       addParser "INVITE", 2, 2, "parserInvitationReceived"
          
770       addParser RPL_INVITING, 3, 3, "parserInviteUser"
End Sub

Public Sub parserPing(origin As COrigin, command As String, paramCount As Byte, params() As String)
10        sendLine "PONG :" & params(0)
End Sub

Public Sub parserError(origin As COrigin, command As String, paramCount As Byte, params() As String)
10        m_statusWindowText.addEvent "IRC_ERROR", makeStringArray(params(0))
End Sub

Public Sub parserKill(origin As COrigin, command As String, paramCount As Byte, params() As String)
10        m_expectingDisconnect = True
20        m_statusWindowText.addEvent "KILLED", makeStringArray(origin.nickname, params(1))
End Sub

Public Sub parserNotice(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        processMessage origin, True, paramCount, params
End Sub

Public Sub parserWelcome(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_registered = True

20        m_currentNickname = params(0)
30        m_serverName = origin.nickname
40        m_statusWindow.switchbartab.caption = m_serverName
50        m_client.switchbar.redrawTab m_statusWindow.switchbartab

60        m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
70        m_statusWindowText.addEvent "WELCOME", makeStringArray(origin.nickname, params(0))
          
80        registered
End Sub

Public Sub parserNicknameInUse(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
10        If Not m_registered Then
20            If m_triedBackupNickname Then
30                m_statusWindowText.addEvent "NICKNAME_IN_USE_PREREG2", makeStringArray(params(1))
40                m_statusWindowText.inputText = "/nick "
50                Beep
60            Else
70                m_statusWindowText.addEvent "NICKNAME_IN_USE_PREREG", makeStringArray(params(1))
80                sendLine "NICK " & m_backupNickname
90                m_triedBackupNickname = True
100           End If
110       Else
120           m_statusWindowText.addEvent "NICKNAME_IN_USE", makeStringArray(params(1))
130       End If
End Sub

Public Sub parserNoSuchNick(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim activeText As ITextWindow
          
10        Set activeText = m_client.activeTextWindow
          
20        If Not activeText Is Nothing Then
30            If activeText.session Is Me Then
40                activeText.addEvent "NO_SUCH_NICK", makeStringArray(params(1))
50            End If
60        End If
End Sub

Public Sub parserISupport(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        parseISupport params(1)
          
20        If m_prefixModes.count = 0 Then
30            parsePrefixModes getISupportValue("PREFIX")
40        End If
          
50        If LenB(m_networkName) = 0 Then
60            m_networkName = getISupportValue("NETWORK")
              
70            If LenB(m_networkName) <> 0 Then
80                m_statusWindow.switchbartab.caption = m_networkName
90                m_client.switchbar.redrawTab m_statusWindow.switchbartab
                  
100               If settings.setting("enableLogging", estBoolean) Then
110                   If settings.setting("logDirectories", estBoolean) Then
120                       If Not settings.setting("logDirectoriesProfile", estBoolean) Or _
                              m_serverProfile Is Nothing Or m_serverProfile Is serverProfiles.quickConnectProfile Then
                              
130                           m_baseLogPath = g_userPath & LOG_DIR & sanitizeFilename(m_networkName) & "\"
                              
140                           If Dir(m_baseLogPath, vbDirectory) = vbNullString Then
150                               MkDir m_baseLogPath
160                           End If
                              
170                           If settings.setting("logStatus", estBoolean) Then
180                               m_statusWindowText.textview.logName = m_baseLogPath & "status"
190                               m_statusWindowText.textview.enableLogging = True
200                               m_statusWindowText.textview.writeEntireBuffer
210                           End If
220                       End If
230                   End If
240               End If
250           End If
260       End If
          
270       If m_channelModes.count = 0 Then
280           parseChannelModes getISupportValue("CHANMODES")
290       End If
          
300       If m_maxModes = 0 Then
310           m_maxModes = CByte(Val(getISupportValue("MODES")))
320       End If
          
330       If LenB(m_channelTypes) = 0 Then
340           m_channelTypes = getISupportValue("CHANTYPES")
350       End If
End Sub

Public Sub parserJoin(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        If StrComp(origin.nickname, m_currentNickname, vbTextCompare) = 0 Then
20            Set channel = findChannel(params(0))
              
30            If channel Is Nothing Then
40                Set channel = createChannel(params(0))
50                m_client.ShowWindow channel.window
60                channel.meJoined
70            Else
80                channel.reJoined
90            End If
100       Else
110           Set channel = findChannel(params(0))
              
120           If Not channel Is Nothing Then
130               channel.userJoined origin
140           End If
150       End If
End Sub

Public Sub parserPart(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel

10        Set channel = findChannel(params(0))
          
20        If channel Is Nothing Then
30            Exit Sub
40        End If
          
50        If StrComp(origin.nickname, m_currentNickname, vbTextCompare) = 0 Then
60            If channel.hopping Then
70                rejoinChannel channel
80            Else
                  Dim count As Long
                  
90                For count = 1 To m_channelsToDestroy.count
100                   If m_channelsToDestroy.item(count) Is channel Then
110                       m_channelsToDestroy.Remove count
120                       Exit For
130                   End If
140               Next count
                  
150               destroyChannel channel
160           End If
170       Else
180           If paramCount > 1 Then
190               channel.userParted origin, params(1)
200           Else
210               channel.userParted origin
220           End If
230       End If
End Sub

Public Sub parserKick(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(0))
          
20        If channel Is Nothing Then
30           Exit Sub
40        End If
          
50        If StrComp(params(1), m_currentNickname, vbTextCompare) = 0 Then
60            channel.meKicked origin, params(2)
70        Else
80            channel.userKicked origin, params(1), params(2)
90        End If
End Sub

Public Sub parserQuit(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        For Each channel In m_channels
20            If channel.isIn(origin.nickname) Then
30                If paramCount = 0 Then
40                    channel.userQuit origin
50                Else
60                    channel.userQuit origin, params(0)
70                End If
80            End If
90        Next channel
End Sub

Public Sub parserNick(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        If StrComp(origin.nickname, m_currentNickname, vbTextCompare) = 0 Then
20            m_currentNickname = params(0)
30            m_statusWindowText.addEvent "ME_NICK_CHANGE", makeStringArray(params(0))
40        End If
          
50        For Each channel In m_channels
60            If channel.isIn(origin.nickname) Then
70                channel.userNickChange origin, params(0)
80            End If
90        Next channel
          
          
100       If Not findQuery(params(0)) Is Nothing Then
110           Exit Sub
120       End If
          
          Dim query As CQuery
          
130       Set query = findQuery(origin.nickname)
          
140       If Not query Is Nothing Then
150           query.nickChange origin, params(0)
160           m_queries.Remove LCase$(origin.nickname)
170           m_queries.Add query, LCase$(query.name)
180       End If
End Sub

Public Sub parserNames(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(2))
          
20        If Not channel Is Nothing Then
30            channel.processNamesList params(3)
40        End If
End Sub

Public Sub parserEndOfNames(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.endOfNames
40        End If
End Sub

Public Sub parserWhoisUser(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        Set m_whoisText = m_client.activeTextWindow
          
20        If Not m_whoisText.session Is Me Then
30            Set m_whoisText = m_statusWindow
40        End If
          
50        If Not m_whoisText Is Nothing Then
60            m_whoisText.addEvent "WHOIS_USER", makeStringArray(params(1), params(2), params(3), _
                  params(5))
70        End If
End Sub

Public Sub parserWhoisChannels(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_CHANNELS", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisServer(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_SERVER", makeStringArray(params(1), params(2), params(3))
30        End If
End Sub

Public Sub parserWhoisRegNick(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_REGNICK", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisOperator(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_OPERATOR", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisHelpOp(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_OPERATOR", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisSecure(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisBot(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisFilter(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisFilterOver(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_GENERIC", makeStringArray(params(1), params(2))
30        End If
End Sub

Public Sub parserWhoisIdle(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_IDLE", makeStringArray(params(1), timeString(Val(params(2))), _
                  formatTime(Val(params(3))))
30        End If
End Sub

Public Sub parserEndOfWhois(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "WHOIS_END", makeStringArray(params(1), params(2))
30            Set m_whoisText = Nothing
40        End If
End Sub

Public Sub parserAway(origin As COrigin, command As String, paramCount As Byte, params() As String)
10        If Not m_whoisText Is Nothing Then
20            m_whoisText.addEvent "AWAY", makeStringArray(params(1), params(2))
30        Else
              Dim activeText As ITextWindow
              
40            Set activeText = m_client.activeTextWindow
              
50            If Not activeText.session Is Me Then
60                Set activeText = m_statusWindow
70            End If
              
80            activeText.addEvent "AWAY", makeStringArray(params(1), params(2))
90        End If
End Sub

Public Sub parserBannedFromChan(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.IText.addEvent "ERR_BANNEDFROMCHAN", makeStringArray(params(1))
40        End If
          
50        m_statusWindowText.addEvent "ERR_BANNEDFROMCHAN", makeStringArray(params(1))
End Sub

Public Sub parserNeedReggedNick(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.IText.addEvent "ERR_NEEDREGGEDNICK", makeStringArray(params(1))
40        End If
          
50        m_statusWindowText.addEvent "ERR_NEEDREGGEDNICK", makeStringArray(params(1))
End Sub

Public Sub parserInviteOnlyChan(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.IText.addEvent "ERR_INVITEONLYCHAN", makeStringArray(params(1))
40        End If
          
50        m_statusWindowText.addEvent "ERR_INVITEONLYCHAN", makeStringArray(params(1))
End Sub

Public Sub parserBadChannelKey(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.IText.addEvent "ERR_BADCHANNELKEY", makeStringArray(params(1))
40        End If
          
50        m_statusWindowText.addEvent "ERR_BADCHANNELKEY", makeStringArray(params(1))
End Sub

Public Sub parserChannelIsFull(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.IText.addEvent "ERR_CHANNELISFULL", makeStringArray(params(1))
40        End If
          
50        m_statusWindowText.addEvent "ERR_CHANNELISFULL", makeStringArray(params(1))
End Sub

Public Sub parserTooManyJoins(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.IText.addEvent "ERR_TOOMANYJOINS", makeStringArray(params(1))
40        End If
          
50        m_statusWindowText.addEvent "ERR_TOOMANYJOINS", makeStringArray(params(1))
End Sub

Public Sub parserTooManyChannels(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
10        m_statusWindowText.addEvent "ERR_TOOMANYCHANNELS", makeStringArray(params(1))
End Sub

Public Sub parserSecureOnlyChan(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)

10        m_statusWindowText.addEvent "ERR_SECUREONLYCHAN", makeStringArray(params(1))
End Sub

Public Sub parserLinkChannel(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
End Sub

Public Sub parserNoPrivileges(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "ERR_NOPRIVILEGES", makeStringArray(params(1))
End Sub

Public Sub parserOperOnly(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
End Sub

Public Sub parserUnknownCommand(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
10        m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1) & " " & params(2))
End Sub

Public Sub parserErroneusNickname(origin As COrigin, command As String, paramCount As Byte, _
    params() As String)
10        m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1) & " " & params(2))
End Sub

Public Sub parserMode(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(0))

20        If Not channel Is Nothing Then
30            If paramCount > 2 Then
40                channel.modeChange origin, params(1), params(2)
50            Else
60                channel.modeChange origin, params(1), vbNullString
70            End If
80        Else
90            If StrComp(params(0), m_currentNickname, vbTextCompare) = 0 Then
100               m_statusWindowText.addEvent "ME_MODE_CHANGE", makeStringArray(origin.nickname, params(1))
110           End If
120       End If
End Sub

Public Sub parserChannelModeIs(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            If paramCount > 3 Then
40                channel.modeIs params(2), params(3)
50            Else
60                channel.modeIs params(2)
70            End If
80        End If
End Sub

Public Sub parserTopicIs(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.topicIs params(2)
40        End If
End Sub

Public Sub parserTopicWhoTime(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.topicWhoTime params(2), Val(params(3))
40        End If
End Sub

Public Sub parserTopic(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(0))
          
20        If Not channel Is Nothing Then
30            channel.topicChange origin, params(1)
40        End If
End Sub

Public Sub parserPrivmsg(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          
10        processMessage origin, False, paramCount, params
End Sub

Public Sub parserBanList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.modeListEntry "b", params(2)
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "BAN_LIST", makeStringArray(params(1), params(2), params(3), _
                  formatTime(params(4)))
70        End If
End Sub

Public Sub parserEndOfBanList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.endOfModeList "b"
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "END_OF_BAN_LIST", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserExList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.modeListEntry "e", params(2)
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "EX_LIST", makeStringArray(params(1), params(2), params(3), _
                  formatTime(params(4)))
70        End If
End Sub

Public Sub parserEndOfExList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.endOfModeList "e"
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "END_OF_EX_LIST", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserInvexList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.modeListEntry "I", params(2)
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "INVEX_LIST", makeStringArray(params(1), params(2), params(3), _
                  formatTime(params(4)))
70        End If
End Sub

Public Sub parserEndOfInvexList(origin As COrigin, command As String, paramCount As Byte, params() _
    As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.endOfModeList "I"
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "END_OF_INVEX_LIST", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserAList(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.modeListEntry "a", params(2)
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "A_LIST", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserEndOfAList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.endOfModeList "a"
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "END_OF_A_LIST", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserQList(origin As COrigin, command As String, paramCount As Byte, params() As String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.modeListEntry "q", params(2)
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "Q_LIST", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserEndOfQList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        Set channel = findChannel(params(1))
          
20        If Not channel Is Nothing Then
30            channel.endOfModeList "q"
40        End If
          
50        If Not CCOpen() Then
60            m_statusWindowText.addEvent "END_OF_Q_LIST", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserWhoList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        Set m_whoisText = m_client.activeTextWindow
          
20        If Not m_whoisText.session Is Me Then
30            Set m_whoisText = m_statusWindow
40        End If
          
50        If Not m_whoisText Is Nothing Then
60            m_whoisText.addEvent "WHO_LIST", makeStringArray(params(1), params(5), params(2), _
                  params(3), params(6), params(7))
70        End If
End Sub

Public Sub parserEndOfWhoList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        Set m_whoisText = m_client.activeTextWindow
          
20        If Not m_whoisText.session Is Me Then
30            Set m_whoisText = m_statusWindow
40        End If
          
50        If Not m_whoisText Is Nothing Then
60            m_whoisText.addEvent "WHO_END_OF_LIST", makeStringArray(params(2))
70        End If
End Sub

Public Sub parserWhowasHost(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        Set m_whoisText = m_client.activeTextWindow
          
20        If Not m_whoisText.session Is Me Then
30            Set m_whoisText = m_statusWindow
40        End If
          
50        If Not m_whoisText Is Nothing Then
60            m_whoisText.addEvent "WHOWAS_HOST", makeStringArray(params(1), params(2), params(3), params(5))
70        End If
End Sub

Public Sub parserWhowasUnknown(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        Set m_whoisText = m_client.activeTextWindow
          
20        If Not m_whoisText.session Is Me Then
30            Set m_whoisText = m_statusWindow
40        End If
          
50        If Not m_whoisText Is Nothing Then
60            m_whoisText.addEvent "WHOWAS_UNKNOWN", makeStringArray(params(1), params(2))
70        End If
End Sub

Public Sub parserEndOfWhowas(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        Set m_whoisText = m_client.activeTextWindow
          
20        If Not m_whoisText.session Is Me Then
30            Set m_whoisText = m_statusWindow
40        End If
          
50        If Not m_whoisText Is Nothing Then
60            m_whoisText.addEvent "WHOWAS_END", makeStringArray(params(2))
70        End If
End Sub

Public Sub parserSilenceModify(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "SILENCE_MODIFY", makeStringArray(params(0))
End Sub

Public Sub parserInvitationReceived(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          
10        If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_INVITE) Then
20            Exit Sub
30        End If
          
40        If m_client.activeTextWindow.session Is Me Then
50            m_client.activeTextWindow.addEvent "INVITATION_RECEIVED", makeStringArray(params(1), _
                  origin.nickname, origin.username, origin.hostname)
60        Else
70            m_statusWindowText.addEvent "INVITATION_RECEIVED", makeStringArray(params(1), _
                  origin.nickname, origin.username, origin.hostname)
80        End If
End Sub

Public Sub parserListStart(rigin As COrigin, command As String, paramCount As Byte, params() As _
    String)

10        If m_channelListWindow Is Nothing Then
20            showList
30        Else
40            m_channelListWindow.clear
50            m_client.ShowWindow m_channelListWindow
60        End If
End Sub

Public Sub parserList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
          
10        If Not m_channelListWindow Is Nothing Then
20            m_channelListWindow.insertChannel params(1), Val(params(2)), params(3)
30        End If
End Sub

Public Sub parserListEnd(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_listSynced = True
20        m_statusWindowText.addEvent "NUMERIC", makeStringArray(params(1))
End Sub

Public Sub parserSilenceList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "SILENCE_LIST", makeStringArray(params(2))
End Sub

Public Sub parserEndOfSilenceList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "END_OF_SILENCE_LIST", makeStringArray(params(1))
End Sub

Public Sub parserChgNickModeN(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "ERR_CHGNICK_MODEN", makeStringArray(params(1))
End Sub

Public Sub parserChgNickModeB(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "ERR_CHGNICK_MODEB", makeStringArray(params(1), params(2))
End Sub

Public Sub parserHelpopTitle(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "HELPOP_TITLE", makeStringArray(params(1))
End Sub

Public Sub parserHelpopText(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "HELPOP_TEXT", makeStringArray(params(1))
End Sub

Public Sub parserMarkedAway(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "MARKED_AWAY", makeStringArray(params(1))
End Sub

Public Sub parserUnmarkedAway(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "NO_LONGER_MARKED_AWAY", makeStringArray(params(1))
End Sub

Public Sub parserInviteList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "INVITE_LIST", makeStringArray(params(1))
End Sub

Public Sub parserEndOfInviteList(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "END_OF_INVITE_LIST", makeStringArray(params(1))
End Sub

Public Sub parserInviteUser(origin As COrigin, command As String, paramCount As Byte, params() As _
    String)
10        m_statusWindowText.addEvent "INVITE_USER", makeStringArray(params(1), params(2))
End Sub

Private Sub processMessage(origin As COrigin, notice As Boolean, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          Dim query As CQuery
          Dim prefix As String
          Dim channelName As String
          Dim activeText As ITextWindow
          
10        On Error GoTo processMessage_Error

20        Set activeText = m_client.activeTextWindow
          
30        If Not activeText Is Nothing Then
40            If Not activeText.session Is Me Then
50                Set activeText = m_statusWindow
60            End If
70        Else
80            Set activeText = m_statusWindow
90        End If
          
100       channelName = params(0)
          
110       If getPrefixIndex(left$(channelName, 1)) <> 0 Then
120           prefix = left$(channelName, 1)
130           channelName = Mid$(channelName, 2)
140       End If

150       Set channel = findChannel(channelName)
          
160       If channel Is Nothing Then
170           If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_PRIVATE) Then
180               Exit Sub
190           End If
          
200           Set query = findQuery(origin.nickname)
210       Else
220           If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_CHANNEL) Then
230               Exit Sub
240           End If
250       End If
          
          Dim text As String
          Dim highlight As Boolean
          
260       text = params(1)
          
270       If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_CODES) Then
280           text = stripFormattingCodes(text)
290       End If
          
300       If settings.enableHighlighting Then
310           If settings.highlightNickname Then
320               If InStr(1, text, m_currentNickname, vbTextCompare) Then
330                   highlight = True
340               End If
350           End If
              
360           If Not highlight Then
                  Dim count As Long
                  
370               For count = 1 To highlights.highlightCount
380                   If InStr(1, text, highlights.highlightItem(count).text, vbTextCompare) Then
390                       highlight = True
400                       Exit For
410                   End If
420               Next count
430           End If
440       End If
          
450       If highlight Then
460           highlighted
470       End If
          
480       If left$(text, 1) = Chr$(1) Then
              Dim ctcpCommand As String
              Dim ctcpParam As String
              
490           extractCtcp text, ctcpCommand, ctcpParam
              
500           If StrComp(ctcpCommand, "ACTION", vbTextCompare) = 0 Then
510               If Not channel Is Nothing Then
520                   channel.userEmote origin, ctcpParam, prefix, highlight
530               Else
540                   If InStr(m_channelTypes, left$(channelName, 1)) <> 0 Then
550                       Exit Sub
560                   End If
                  
570                   If query Is Nothing Then
580                       Set query = createQuery(origin.nickname)
590                   End If
                      
600                   query.userEmote origin, ctcpParam, highlight
610               End If
620           Else
630               If ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_CTCP) Then
640                   Exit Sub
650               End If
              
660               If notice Then
670                   processCtcpReply origin, ctcpCommand, ctcpParam
680               Else
690                   processCtcp origin, params(0), ctcpCommand, ctcpParam
700               End If
710           End If
720       Else
              'Handled after ctcp processing to avoid unintentionally
              'blocking ctcp replies (which are really just notices)
730           If notice And ignoreManager.isIgnored(origin.nickname & "!" & origin.username & "@" & origin.hostname, IGNORE_NOTICE) Then
740               Exit Sub
750           End If
          
760           If Not channel Is Nothing Then
770               channel.userMessage origin, text, notice, prefix, highlight
780           Else
790               If notice Then
800                   m_statusWindowText.addEvent "NOTICE", makeStringArray(origin.nickname, text)
810                   m_client.switchbar.tabActivity m_statusWindow.switchbartab, tasMessage
                      
820                   If Not activeText Is Nothing And Not activeText Is m_statusWindow Then
830                       activeText.addEvent "NOTICE", makeStringArray(origin.nickname, text)
                          
                          Dim tabWindow As ITabWindow
840                       Set tabWindow = activeText
                          
850                       m_client.switchbar.tabActivity tabWindow.getTab, tasMessage
860                   End If
870               Else
880                   If InStr(m_channelTypes, left$(channelName, 1)) <> 0 Then
890                       Exit Sub
900                   End If
                  
910                   If query Is Nothing Then
920                       Set query = createQuery(origin.nickname)
930                   End If
                      
940                   query.userMessage origin, text, highlight
950               End If
960           End If
970       End If

980      On Error GoTo 0
990      Exit Sub

processMessage_Error:
1000      handleError "processMessage", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub highlighted()
10        If settings.setting("highlightCustomSound", estBoolean) Then
20            PlaySound settings.setting("highlightSoundPath", estString)
30        Else
40            Beep
50        End If
End Sub

Private Sub PlaySound(soundPath As String)
          Dim command As String
          Dim alias As String
          Dim path As String
          Dim ext As String
          Dim audioType As String
          
10        Randomize
20        alias = "MySound" & Rnd
30        path = soundPath
40        ext = right$(path, 4)
          
50        If ext = ".mp3" Then
60            audioType = "mpegvideo"
70        ElseIf ext = ".wav" Then
80            audioType = "waveaudio"
90        Else
100           audioType = vbNullString
110           Beep
120       End If
          
130       If audioType <> vbNullString Then
140           command = "open """ & path & _
                  """ type " & audioType & _
                  " alias " & alias
                  
150           If mciSendString(command, 0, 0, 0) = 0 Then
160               command = "play " & alias & _
                  " from 0"
170               mciSendString command, 0, 0, 0
180           End If
190       End If
End Sub

Private Sub extractCtcp(text As String, ByRef command As String, ByRef param As String)
          Dim inParam As Boolean
          
10        Do While LenB(text) <> 0
20            If left$(text, 1) = Chr$(1) Then
30                If LenB(command) <> 0 Then
40                    Exit Sub
50                End If
60            ElseIf left$(text, 1) = " " Then
70                If LenB(param) = 0 Then
80                    inParam = True
90                Else
100                   param = param & left$(text, 1)
110               End If
120           Else
130               If inParam Then
140                   param = param & left$(text, 1)
150               Else
160                   command = command & left$(text, 1)
170               End If
180           End If
              
190           text = Mid$(text, 2)
200       Loop
End Sub

Private Sub processCtcp(origin As COrigin, target As String, command As String, param As String)
10        If GetTickCount - m_lastCtcp < 2500 Then
20            Exit Sub
30        End If
          
          Dim prefix As String
          Dim channelName As String
          Dim channel As CChannel
          
40        If getPrefixIndex(left$(target, 1)) <> 0 Then
50            prefix = left$(target, 1)
60            channelName = Mid$(target, 2)
70        End If

80        Set channel = findChannel(channelName)
          
          Dim activeWindow As ITextWindow
          
90        Set activeWindow = m_client.activeTextWindow
          
100       If Not activeWindow Is Nothing Then
110           If Not activeWindow.session Is Me Then
120               Set activeWindow = m_statusWindow
130           End If
140       Else
150           Set activeWindow = m_statusWindow
160       End If
          
170       Select Case UCase$(command)
              Case "PING"
180               sendCtcpReply origin.nickname, "PING", param
                  
190               If Not channel Is Nothing Then
200                   channel.userCtcp origin, command, vbNullString, prefix
210               Else
220                   activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command)
230               End If
240           Case "VERSION"
250               sendCtcpReply origin.nickname, "VERSION", "SwiftIRC client V" & m_client.getVersion
                  
260               If Not channel Is Nothing Then
270                   channel.userCtcp origin, command, vbNullString, prefix
280               Else
290                   activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command)
300               End If
310           Case "TIME"
320               sendCtcpReply origin.nickname, "TIME", format(Now, "yyyy-mm-dd hh:MM:ss")
                  
330               If Not channel Is Nothing Then
340                   channel.userCtcp origin, "TIME", vbNullString, prefix
350               Else
360                   activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command)
370               End If
380           Case Else
390               If Not channel Is Nothing Then
400                   channel.userCtcp origin, command, param, prefix
410               Else
420                   activeWindow.addEvent "CTCP_RECEIVED", makeStringArray(origin.nickname, command, param)
430               End If
440       End Select
          
450       m_lastCtcp = GetTickCount
End Sub

Private Sub processCtcpReply(origin As COrigin, command As String, param As String)
          Dim ctcpParam As String

10        Select Case UCase$(command)
              Case "PING"
                  Dim since As Long
                  
20                If LenB(param) <> 0 Then
30                    If Val(param) <= 2 ^ 31 Then
40                        since = GetTickCount - Val(param)
50                        ctcpParam = since & "ms"
60                    Else
70                        ctcpParam = param
80                    End If
90                End If
100           Case Else
110               ctcpParam = param
120       End Select

130       If Not m_client.activeTextWindow Is Nothing Then
140           If m_client.activeTextWindow.session Is Me Then
150               m_client.activeTextWindow.addEvent "CTCP_REPLY_RECEIVED", makeStringArray(origin.nickname, command, ctcpParam)
160           Else
170               m_statusWindowText.addEvent "CTCP_REPLY_RECEIVED", makeStringArray(origin.nickname, command, ctcpParam)
180           End If
190       Else
200           m_statusWindowText.addEvent "CTCP_REPLY_RECEIVED", makeStringArray(origin.nickname, command, ctcpParam)
210       End If
End Sub

Public Sub sendCtcp(target As String, ctcp As String)
10        Select Case UCase$(ctcp)
              Case "PING"
20                sendLine "PRIVMSG " & target & " :" & Chr$(1) & "PING " & GetTickCount & Chr$(1)
30            Case Else
40                sendLine "PRIVMSG " & target & " :" & Chr$(1) & ctcp & Chr$(1)
50        End Select
          
60        If Not m_client.activeTextWindow Is Nothing Then
70            If m_client.activeTextWindow.session Is Me Then
80                m_client.activeTextWindow.addEvent "CMD_CTCP_SENT", makeStringArray(target, ctcp)
90            Else
100               m_statusWindowText.addEvent "CMD_CTCP_SENT", makeStringArray(target, ctcp)
110           End If
120       Else
130           m_statusWindowText.addEvent "CMD_CTCP_SENT", makeStringArray(target, ctcp)
140       End If
End Sub

Private Sub sendCtcpReply(target As String, command As String, param As String)
10        sendLine "NOTICE " & target & " :" & Chr$(1) & command & " " & param & Chr$(1)
End Sub

Private Function createChannel(name As String) As CChannel
          Dim channelWindow As ctlWindowChannel
          Dim channel As New CChannel
          
10        Set channelWindow = m_client.createNewWindow("swiftirc.ctlWindowChannel", "channel")
          
20        channel.init name, channelWindow, Me
          
30        channelWindow.init Me, channel
40        channelWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, channelWindow, sboChannel, _
              channel.name, g_iconSBChannel)
          
50        m_channels.Add channel, LCase$(name)
60        Set createChannel = channel
End Function

Public Sub query(name As String, Optional text As String = vbNullString)
          Dim query As CQuery
          
10        Set query = findQuery(name)
          
20        If Not query Is Nothing Then
30            m_client.ShowWindow query.window
40        Else
50            Set query = createQuery(name)
60            m_client.ShowWindow query.window
70        End If
          
80        If LenB(text) <> 0 Then
90            query.textEntered text
100       End If
End Sub

Private Function createQuery(name As String) As CQuery
          Dim queryWindow As ctlWindowQuery
          Dim query As New CQuery
          
10        Set queryWindow = m_client.createNewWindow("swiftIrc.ctlWindowQuery", "query")
20        query.init name, queryWindow, Me
30        queryWindow.init Me, query
40        queryWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, queryWindow, sboQuery, query.name, g_iconSBQuery)

50        m_queries.Add query, LCase$(name)
60        Set createQuery = query
End Function

Private Sub destroyChannel(channel As CChannel)
10        debugLog "Destroying channel " & channel.name

20        m_client.removeTab channel.window.switchbartab
30        m_client.destroyWindow channel.window
40        m_channels.Remove LCase$(channel.name)
50        channel.deInit
End Sub

Private Sub destroyQuery(query As CQuery)
10        debugLog "Destroying query " & query.name

20        m_client.removeTab query.window.switchbartab
30        m_client.destroyWindow query.window
40        m_queries.Remove LCase$(query.name)
50        query.deInit
End Sub

Public Sub partChannel(channel As CChannel)
10        sendLine "PART " & channel.name
20        destroyChannel channel
End Sub

Public Sub closeQuery(query As CQuery)
10        destroyQuery query
End Sub

Public Sub closeGenericWindow(genericWindow As ctlWindowGenericText)
10        m_client.removeTab genericWindow.switchbartab
20        m_client.destroyWindow genericWindow
          
30        If genericWindow Is m_debugWindow Then
40            Set m_debugWindow = Nothing
50            Set m_debugWindowText = Nothing
60        End If
End Sub

Private Function findChannel(name As String) As CChannel
10        On Error Resume Next
20        Set findChannel = m_channels.item(LCase$(name))
End Function

Private Function findQuery(name As String) As CQuery
10        On Error Resume Next
20        Set findQuery = m_queries.item(LCase$(name))
End Function

Friend Sub joinChannel(name As String)
10        sendLine "JOIN " & name
End Sub

Friend Sub rejoinChannel(channel As CChannel)
10        channel.reJoining

20        If LenB(channel.key) <> 0 Then
30            sendLine "JOIN " & channel.name & " " & channel.key
40        Else
50            sendLine "JOIN " & channel.name
60        End If
End Sub

Private Sub addISupport(name As String, value As String)
          Dim iSupport As New CISupportEntry

10        iSupport.init name, value
          
20        On Error Resume Next
30        m_iSupport.Add iSupport, UCase(name)
End Sub

Private Function getISupportValue(name As String) As String
10        On Error Resume Next
          
20        getISupportValue = m_iSupport.item(UCase(name)).value
End Function

Private Sub parseISupport(iSupport As String)
          Dim count As Integer
          Dim char As String
          
          Dim name As String
          Dim value As String
          Dim inValue As Boolean
          
10        If right$(iSupport, 1) <> " " Then
20            iSupport = iSupport & " "
30        End If
          
40        For count = 1 To Len(iSupport)
50            char = Mid$(iSupport, count, 1)
              
60            If char = " " Then
70                If LenB(name) <> 0 Then
80                    If LenB(value) <> 0 Then
90                        addISupport name, value
100                       value = vbNullString
110                   Else
120                       addISupport name, "1"
130                   End If
                      
140                   name = vbNullString
150                   inValue = False
                      
160                   If count < Len(iSupport) Then
170                       If Mid$(iSupport, count + 1, 1) = ":" Then
180                           Exit For
190                       End If
200                   End If
210               End If
220           ElseIf char = "=" And Not inValue Then
230               inValue = True
240           Else
250               If inValue Then
260                   value = value & char
270               Else
280                   name = name & char
290               End If
300           End If
310       Next count
End Sub

Private Sub addPrefixMode(mode As String, symbol As String)
          Dim prefixMode As New CPrefixMode
          
10        prefixMode.init mode, symbol
20        m_prefixModes.Add prefixMode
End Sub

Public Function comparePrefix(prefix1 As String, prefix2 As String) As Integer
          Dim index1 As Byte
          Dim index2 As Byte
          
10        index1 = getPrefixIndex(prefix1)
20        index2 = getPrefixIndex(prefix2)
          
30        If index1 = 0 And index2 = 0 Then
40            comparePrefix = 0
50        ElseIf index1 = 0 Then
60            comparePrefix = 1
70        ElseIf index2 = 0 Then
80            comparePrefix = -1
90        ElseIf index1 = index2 Then
100           comparePrefix = 0
110       ElseIf index1 < index2 Then
120           comparePrefix = -1
130       Else
140           comparePrefix = 1
150       End If
End Function

Friend Function getUserStyle(nickname As String, prefix As String) As CUserStyle
10        If StrComp(nickname, m_currentNickname, vbTextCompare) = 0 _
              And settings.setting("colourMyNick", estBoolean) Then
              
              Dim prefixIndex As Long
20            prefixIndex = getPrefixIndex(prefix)
              
30            If prefixIndex = 0 Then
40                Set getUserStyle = styleMe
50            ElseIf prefixIndex <= getPrefixIndex("@") Then
60                Set getUserStyle = styleMeOp
70            ElseIf prefixIndex <= getPrefixIndex("%") Then
80                Set getUserStyle = styleMeHalfop
90            ElseIf prefixIndex <= getPrefixIndex("+") Then
100               Set getUserStyle = styleMeVoice
110           End If
120       Else
130           Set getUserStyle = getPrefixStyle(prefix)
              
140           If getUserStyle Is Nothing Then
150               Set getUserStyle = styleNormal
160           End If
170       End If
End Function

Friend Function getPrefixStyle(symbol As String) As CUserStyle
          Dim count As Long
          
          Dim index As Long
          
10        index = getPrefixIndex(symbol)
          
20        If index = 0 Then
30            Exit Function
40        End If
          
50        For count = 1 To prefixStyles.count
60            If index <= getPrefixIndex(prefixStyles.item(count).symbol) Then
70                Set getPrefixStyle = prefixStyles.item(count).style
80                Exit Function
90            End If
100       Next count
End Function

Private Function getPrefixIndex(symbol As String) As Byte
          Dim count As Integer
          
10        For count = 1 To m_prefixModes.count
20            If m_prefixModes.item(count).symbol = symbol Then
30                getPrefixIndex = count
40                Exit Function
50            End If
60        Next count
End Function

Private Function getPrefixModeIndex(mode As String) As Byte
          Dim count As Integer
          
10        For count = 1 To m_prefixModes.count
20            If m_prefixModes.item(count).mode = mode Then
30                getPrefixModeIndex = count
40                Exit Function
50            End If
60        Next count
End Function

Public Function isPrefixMode(mode As String) As Boolean
10        isPrefixMode = (getPrefixModeIndex(mode) <> 0)
End Function

Public Function getPrefixModeBySymbol(symbol As String) As String
          Dim prefix As CPrefixMode
          
10        Set prefix = findPrefixBySymbol(symbol)
          
20        If Not prefix Is Nothing Then
30            getPrefixModeBySymbol = prefix.mode
40        End If
End Function

Public Function getEffectivePrefix(modes As String) As String
          Dim count As Integer
          Dim highest As Byte
          Dim index As Byte
          
10        For count = 1 To Len(modes)
20            index = getPrefixModeIndex(Mid$(modes, count, 1))
              
30            If index < highest Or highest = 0 Then
40                highest = index
50            End If
60        Next count
          
70        If highest <> 0 Then
80            getEffectivePrefix = m_prefixModes.item(highest).symbol
90        End If
End Function

Private Function findPrefixBySymbol(symbol As String) As CPrefixMode
          Dim count As Integer
          
10        For count = 1 To m_prefixModes.count
20            If m_prefixModes.item(count).symbol = symbol Then
30                Set findPrefixBySymbol = m_prefixModes.item(count)
40                Exit Function
50            End If
60        Next count
End Function

Private Sub parsePrefixModes(prefixes As String)
          Dim count As Integer
          Dim char As String
          
          Dim modes As String
          Dim symbols As String
          
          Dim inModes As Boolean
          
10        For count = 1 To Len(prefixes)
20            char = Mid$(prefixes, count, 1)
              
30            If char = "(" Then
40                inModes = True
50            ElseIf char = ")" Then
60                inModes = False
70            Else
80                If inModes Then
90                    modes = modes & char
100               Else
110                   symbols = symbols & char
120               End If
130           End If
140       Next count
          
150       For count = 1 To Len(modes)
160           If count <= Len(symbols) Then
170               addPrefixMode Mid$(modes, count, 1), Mid$(symbols, count, 1)
180           End If
190       Next count
End Sub

Private Sub parseChannelModes(modes As String)
          Dim modeType As eChannelModeType
          Dim channelMode As CChannelMode
          Dim count As Long
          Dim mode As String
          
10        For count = 1 To Len(modes)
20            mode = Mid$(modes, count, 1)
              
30            If mode = "," Then
40                If modeType <> cmtUnknown Then
50                    modeType = modeType + 1
60                End If
70            Else
80                Set channelMode = New CChannelMode
90                channelMode.init mode, modeType
100               m_channelModes.Add channelMode
110           End If
120       Next count
End Sub

Public Function getListModes() As String
          Dim count As Long
          
10        For count = 1 To m_channelModes.count
20            If m_channelModes.item(count).modeType = cmtList Then
30                getListModes = getListModes & m_channelModes.item(count).mode
40            End If
50        Next count
End Function

Public Function getChannelModeType(mode As String) As eChannelModeType
          Dim count As Long
          
10        For count = 1 To m_channelModes.count
20            If m_channelModes.item(count).mode = mode Then
30                getChannelModeType = m_channelModes.item(count).modeType
40                Exit Function
50            End If
60        Next count
          
70        getChannelModeType = cmtUnknown
End Function

Private Sub m_socket_closed(errorCode As Long)
10        disconnected
End Sub

Private Sub m_socket_connectFailed(errorCode As Long)
10        Select Case errorCode
              Case 10060
20                m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                      textManager.getText("ERR_CONNECT_TIMEOUT"))
30            Case 10061
40                m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                      textManager.getText("ERR_CONNECT_REFUSED"))
50            Case 10053
60                m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                      textManager.getText("ERR_CONNECT_ABORTED"))
70            Case 10055
80                m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                      textManager.getText("ERR_CONNECT_BUFSPACE"))
90            Case 8
100               m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
                      textManager.getText("ERR_CONNECT_MEMORY"))
110           Case Else
120               m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, "Code: " _
                      & errorCode)
130       End Select
          
140       connectFailed
End Sub

Private Sub m_socket_resolveCompleteA(requestId As Long, addrList As cArrayList)
10        If requestId = m_resolveIdConnect Then
20            Set m_addrList = addrList
30            m_currentAddrIndex = 1
40            m_socket.connect addrList.item(m_currentAddrIndex).addr, m_serverPort
50            m_resolveIdConnect = 0
60        End If
End Sub

Private Sub m_socket_resolveFailed(requestId As Long, errorCode As Long)
10        m_resolveIdConnect = 0
20        m_statusWindowText.addEvent "ERROR_CONNECT", makeStringArray(m_serverHost, _
              textManager.getText("ERR_CONNECT_DNS"))
              
30        connectFailed
End Sub

Private Sub m_statusWindow_textInput(text As String)
10        textInput m_statusWindow, text
End Sub

Public Sub textInput(Source As ITextWindow, line As String)
10        line = LTrim(line)
          
20        If left$(line, 1) <> "/" Then
30            Exit Sub
40        End If
          
          Dim command As String
          
50        extractCommand line, command
          
          Dim inputParser As CInputParser
          
60        Set inputParser = findInputParser(command)
          
70        If inputParser Is Nothing Then
80            sendLine command & " " & line
90            Exit Sub
100       End If
          
          Dim params() As String
          Dim paramCount As Byte
          
110       paramCount = extractCommandParams(line, inputParser.maxParams, params)
          
120       If paramCount >= inputParser.minParams Then
130           CallByName Me, inputParser.func, VbMethod, Source, command, paramCount, params
140       Else
150           Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
160       End If
End Sub

Public Sub commandList(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          
10        If paramCount = 0 Then
20            sendLine "LIST"
30        ElseIf paramCount = 1 Then
40            sendLine "LIST *" & params(0) & "*"
50        Else
60            sendLine "LIST " & params(0) & " " & params(1)
70        End If
End Sub

Public Sub commandJoin(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          
          Dim channelName As String
          Dim key As String
          
10        channelName = params(0)
          
20        If InStr(m_channelTypes, left$(channelName, 1)) = 0 Then
30            channelName = "#" & channelName
40        End If
          
50        If paramCount > 1 Then
60            key = params(1)
70        End If
          
80        If LenB(key) <> 0 Then
              'Key was provided, no need to see
              'if we have a cached one for the channel.
              
90            sendLine "JOIN " & channelName & " " & key
100       Else
              Dim channel As CChannel
              
110           Set channel = findChannel(channelName)
              
120           If Not channel Is Nothing Then
130               If Not channel.meIsIn Then
140                   rejoinChannel channel
150               End If
160           Else
170               sendLine "JOIN " & channelName
180           End If
190       End If
End Sub

Public Sub commandPart(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        If paramCount = 0 Then
20            If TypeOf Source Is ctlWindowChannel Then
                  Dim windowChannel As ctlWindowChannel
30                Set windowChannel = Source
40                partChannel windowChannel.channel
50            Else
60                Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
70            End If
80        ElseIf paramCount = 1 Then
90            sendLine "PART " & params(0)
100       Else
110           sendLine "PART " & params(0) & " : " & params(1)
120       End If
End Sub

Public Sub commandMsg(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        sendLine "PRIVMSG " & params(0) & " :" & params(1)
20        Source.addEvent "CMD_PRIVMSG_SENT", makeStringArray(params(0), params(1))
End Sub

Public Sub commandNotice(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        sendLine "NOTICE " & params(0) & " :" & params(1)
20        Source.addEvent "CMD_NOTICE_SENT", makeStringArray(params(0), params(1))
End Sub

Public Sub commandQuery(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        If paramCount = 1 Then
20            query params(0)
30        Else
40            query params(0), params(1)
50        End If
End Sub

Public Sub commandMe(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        If TypeOf Source Is ctlWindowChannel Then
              Dim windowChannel As ctlWindowChannel
              
20            Set windowChannel = Source
30            windowChannel.channel.sendEmote params(0)
40        ElseIf TypeOf Source Is ctlWindowQuery Then
              Dim windowQuery As ctlWindowQuery
              
50            Set windowQuery = Source
60            windowQuery.query.sendEmote params(0)
70        Else
80            Source.addEvent "CMD_INCOMPATIBLE_WINDOW", makeStringArray(command)
90        End If
End Sub

Public Sub commandTest(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        Source.addText params(0)
End Sub

Public Sub commandTopic(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          
          Dim channelWindow As ctlWindowChannel
          
10        If paramCount > 1 Then
20            sendLine "TOPIC " & params(0) & " :" & params(1)
30        ElseIf paramCount > 0 Then
40            If TypeOf Source Is ctlWindowChannel Then
50                Set channelWindow = Source
60                sendLine "TOPIC " & channelWindow.channel.name & " :" & params(0)
70            Else
80                sendLine "TOPIC " & params(0)
90            End If
100       Else
110           If TypeOf Source Is ctlWindowChannel Then
120               Set channelWindow = Source
130               sendLine "TOPIC " & channelWindow.channel.name
140           End If
150       End If
End Sub

Public Sub commandOptions(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        openOptions m_client, Me
End Sub

Public Sub commandDebug(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        If m_debugWindow Is Nothing Then
20            Set m_debugWindow = m_client.createNewWindow("swiftirc.ctlWindowGenericText", "debug")
30            Set m_debugWindowText = m_debugWindow
              
40            m_debugWindow.init Me, "Debug"
50            m_debugWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, m_debugWindow, sboGeneric, "Debug", g_iconSBGeneric)
60            m_client.ShowWindow m_debugWindow
70        End If
End Sub

Public Sub commandEcho(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        Source.addText params(0)
End Sub

Public Sub commandRaw(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
10        Source.addEvent "CMD_RAW_SENT", makeStringArray(params(0))
20        sendLine params(0)
End Sub

Public Sub commandKick(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          Dim channel As CChannel
          
10        If TypeOf Source Is ctlWindowChannel And InStr(m_channelTypes, left$(params(0), 1)) = 0 Then
              Dim channelWindow As ctlWindowChannel
20            Set channelWindow = Source
              
30            If paramCount > 2 Then
40                sendLine "KICK " & channelWindow.channel.name & " " & params(0) & " :" & params(1) & " " & params(2)
50            ElseIf paramCount > 1 Then
60                sendLine "KICK " & channelWindow.channel.name & " " & params(0) & " :" & params(1)
70            Else
80                sendLine "KICK " & channelWindow.channel.name & " " & params(0) & " :No reason specified"
90            End If
100       Else
110           If paramCount > 1 Then
120               If paramCount > 2 Then
130                   sendLine "KICK " & params(0) & " " & params(1) & " :" & params(2)
140               Else
150                   sendLine "KICK " & params(0) & " " & params(1) & " :No reason specified"
160               End If
170           Else
180               Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
190           End If
200       End If
End Sub

Public Sub commandSlap(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          Dim channelWindow As ctlWindowChannel
          
10        If TypeOf Source Is ctlWindowChannel Then
20            Set channelWindow = Source
30            channelWindow.channel.sendEmote "slaps " & params(0) & " around a bit with a large trout"
40        Else
50            Source.addEvent "CMD_INCOMPATIBLE_WINDOW", makeStringArray(command)
60        End If
End Sub

Public Sub commandHuggle(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          Dim channelWindow As ctlWindowChannel
          
10        If TypeOf Source Is ctlWindowChannel Then
20            Set channelWindow = Source
30            channelWindow.channel.sendEmote "huggles " & params(0)
40        Else
50            Source.addEvent "CMD_INCOMPATIBLE_WINDOW", makeStringArray(command)
60        End If
End Sub

Public Sub commandHop(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          
          Dim channelWindow As ctlWindowChannel
          
10        If TypeOf Source Is ctlWindowChannel Then
20            Set channelWindow = Source
30            channelWindow.channel.hopping = True
              
40            If paramCount < 1 Then
50                channelWindow.channel.hop
60            Else
70                channelWindow.channel.hop params(0)
80            End If
90        Else
100           If paramCount < 1 Then
110               Source.addEvent "CMD_INSUFFICIENT_PARAMS", makeStringArray(command)
120           Else
                  Dim channel As CChannel
                  
130               Set channel = findChannel(params(0))
              
140               If Not channel Is Nothing Then
150                   If paramCount < 2 Then
160                       channelWindow.channel.hop
170                   Else
180                       channelWindow.channel.hop params(1)
190                   End If
200               End If
210           End If
220       End If

End Sub

Private Sub showList()
10        Set m_channelListWindow = m_client.createNewWindow("swiftirc.ctlChannelList", "list")
20        m_channelListWindow.session = Me
30        m_channelListWindow.switchbartab = m_client.switchbar.addTab(m_statusWindow, m_channelListWindow, sboChannelList, "Channels", g_iconSBList)
40        m_client.ShowWindow m_channelListWindow
End Sub

Friend Sub closeChannelList()
10        If Not m_channelListWindow Is Nothing Then
20            m_client.removeTab m_channelListWindow.switchbartab
30            m_client.destroyWindow m_channelListWindow
40            Set m_channelListWindow = Nothing
50        End If
End Sub

Public Sub commandClear(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
          
10        Source.clear
End Sub

Public Sub commandClearAll(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
          
          Dim textWindow As ITextWindow
          Dim channel As CChannel
          Dim query As CQuery
          
10        m_statusWindowText.clear
          
20        For Each channel In m_channels
30            Set textWindow = channel.window
40            textWindow.clear
50        Next channel
          
60        For Each query In m_queries
70            Set textWindow = query.window
80            textWindow.clear
90        Next query
          
100       If Not m_debugWindowText Is Nothing Then
110           m_debugWindowText.clear
120       End If
End Sub

Public Sub commandCtcp(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
10        sendCtcp params(0), params(1)
End Sub

Public Sub commandClose(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
10        If TypeOf Source Is ctlWindowChannel Then
              Dim channelWindow As ctlWindowChannel
              
20            Set channelWindow = Source
30            partChannel channelWindow.channel
40        ElseIf TypeOf Source Is ctlWindowQuery Then
              Dim queryWindow As ctlWindowQuery
          
50            Set queryWindow = Source
60            closeQuery queryWindow.query
70        ElseIf TypeOf Source Is ctlWindowGenericText Then
80            closeGenericWindow Source
90        End If
End Sub

Public Sub commandCloseAll(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
          
10        If TypeOf Source Is ctlWindowChannel Then
              Dim channel As CChannel
              
20            For Each channel In m_channels
30                partChannel channel
40            Next channel
50        ElseIf TypeOf Source Is ctlWindowQuery Then
              Dim query As CQuery
              
60            For Each query In m_queries
70                closeQuery query
80            Next query
90        End If
End Sub


Public Sub commandReconnect(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
10        connect
End Sub

Public Sub commandQuit(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
          
10        m_expectingDisconnect = True
          
20        If paramCount > 0 Then
30            sendLine "QUIT :" & params(0)
40        Else
50            sendLine "QUIT"
60        End If
End Sub

Public Sub commandIgnore(Source As ITextWindow, command As String, paramCount As Byte, params() _
    As String)
          
          Dim commandFlags As String
          Dim ignoreFlags As String
          Dim mask As String
          
10        If paramCount = 0 Then
20            m_client.showIgnoreList Me
30            Exit Sub
40        ElseIf paramCount > 0 Then
50            If left$(params(0), 1) = "-" Then
60                commandFlags = Mid$(params(0), 2)
                  
70                If paramCount > 1 Then
80                    mask = params(1)
                      
90                    If paramCount > 2 Then
100                       ignoreFlags = params(2)
110                   End If
120               End If
130           Else
140               mask = params(0)
                  
150               If paramCount > 1 Then
160                   ignoreFlags = params(1)
170               End If
180           End If
190       End If
          
200       processIgnoreCommand Source, commandFlags, mask, ignoreFlags
End Sub

Private Sub processIgnoreCommand(Source As ITextWindow, commandFlags As String, mask As String, ignoreFlags As String)
          Dim ignore As CIgnoreItem
          Dim removed As Boolean

10        commandFlags = LCase$(commandFlags)

20        If commandFlags = "a" Or commandFlags = "add" Or LenB(commandFlags) = 0 Then
30            If LenB(mask) <> 0 Then
40                Set ignore = New CIgnoreItem
                  
50                If swiftMatch("*!*@*", mask) Then
60                    ignore.mask = mask
70                Else
80                    ignore.mask = getIal(mask, ialHost)
90                End If
                  
100               If LenB(ignoreFlags) <> 0 Then
110                   If Not ignore.parseFlagChars(ignoreFlags) Then
120                       Source.addEvent "IGNORE_INVALID_FLAGS", makeStringArray(ignoreFlags)
130                       Exit Sub
140                   End If
150               Else
160                   ignore.flags = IGNORE_ALL
170               End If
                  
180               If swiftMatch("*!*@*", mask) Then
190                   removed = ignoreManager.removeIgnoreByMask(mask)
200               Else
210                   removed = ignoreManager.removeIgnoreByMask(getIal(mask, ialAll))
220               End If
                  
230               ignoreManager.addIgnore ignore
                  
240               If removed Then
250                   Source.addEvent "IGNORE_UPDATED", makeStringArray(ignore.mask, ignore.flagChars)
260               Else
270                   Source.addEvent "IGNORE_ADDED", makeStringArray(ignore.mask, ignore.flagChars)
280               End If
                  
290               saveIgnoreFile
300           End If
310       ElseIf commandFlags = "l" Or commandFlags = "list" Then
320           listIgnores Source
330       ElseIf commandFlags = "r" Or commandFlags = "remove" Then
340           If LenB(mask) <> 0 Then
350               If swiftMatch("*!*@*", mask) Then
360                   Set ignore = ignoreManager.getIgnoreByMask(mask)
370                   removed = ignoreManager.removeIgnoreByMask(mask)
380               Else
390                   Set ignore = ignoreManager.getIgnoreByMask(getIal(mask, ialAll))
400                   removed = ignoreManager.removeIgnoreByMask(getIal(mask, ialAll))
410               End If
                  
420               If removed Then
430                   Source.addEvent "IGNORE_REMOVED", makeStringArray(ignore.mask)
440                   saveIgnoreFile
450               Else
460                   Source.addEvent "IGNORE_REMOVE_NOTFOUND", makeStringArray(mask)
470               End If
480           End If
490       ElseIf commandFlags = "c" Or commandFlags = "clear" Then
500           ignoreManager.clearIgnores
510           Source.addEvent "IGNORE_LIST_CLEARED", makeStringArray("")
520           saveIgnoreFile
530       Else
540           Source.addEvent "IGNORE_INVALID_COMMAND", makeStringArray(commandFlags)
550       End If
End Sub

Private Sub listIgnores(window As ITextWindow)
          Dim count As Long

10        window.addEvent "IGNORE_LIST_START", makeStringArray("")
          
20        For count = 1 To ignoreManager.ignoreCount
30            window.addEvent "IGNORE_LIST_ENTRY", _
                  makeStringArray(ignoreManager.ignore(count).mask, ignoreManager.ignore(count).flagChars)
40        Next count
          
50        window.addEvent "IGNORE_LIST_END", makeStringArray("")
End Sub


Public Sub commandServer(Source As ITextWindow, command As String, paramCount As Byte, params() As _
    String)
          Dim flags As String
          Dim server As String
          Dim port As Long
          Dim password As String
          
10        If paramCount = 0 Then
20            processServerCommand Source, flags, server, port, password
30            Exit Sub
40        End If

50        If left$(params(0), 1) = "-" Then
60            If Len(params(0)) > 1 Then
70                flags = Mid$(params(0), 2)
80            End If
90        Else
100           If InStr(params(0), ":") <> 0 Then
110               splitServerPort params(0), server, port
120           Else
130               server = params(0)
140           End If
150       End If
          
160       If paramCount > 1 Then
170           If LenB(server) <> 0 Then
180               If port <> 0 Then
190                   password = params(1)
200               Else
210                   port = Val(params(1))
220               End If
230           Else
240               If InStr(params(1), ":") <> 0 Then
250                   splitServerPort params(1), server, port
260               Else
270                   server = params(1)
280               End If
290           End If
300       Else
310           processServerCommand Source, flags, server, port, password
320           Exit Sub
330       End If
          
340       If paramCount > 2 Then
350           If port <> 0 Then
360               password = params(2)
370           Else
380               port = Val(params(2))
390           End If
400       Else
410           processServerCommand Source, flags, server, port, password
420           Exit Sub
430       End If
          
440       If paramCount > 3 Then
450           If LenB(password) <> 0 Then
460               password = params(3)
470           End If
480       End If
          
490       processServerCommand Source, flags, server, port, password
End Sub

Private Sub processServerCommand(Source As ITextWindow, flags As String, server As String, port As _
    Long, password As String)
          
          Dim profile As CServerProfile
          
10        If InStr(flags, "m") <> 0 Then
              Dim newSession As CSession
              
20            Set newSession = m_client.newSession
              
30            If LenB(server) <> 0 Then
40                Set profile = serverProfiles.findProfile(server)
                  
50                If Not profile Is Nothing Then
60                    newSession.serverProfile = profile
70                Else
80                    newSession.serverHost = server
90                    newSession.serverPort = m_serverPort
100                   newSession.primaryNickname = m_primaryNickname
110                   newSession.backupNickname = m_backupNickname
120                   newSession.realName = m_realName
130               End If
                  
140               If port <> 0 Then
150                   newSession.serverPort = port
160               End If
                  
170               If LenB(password) <> 0 Then
180                   newSession.serverPassword = password
190               End If
                  
200               m_client.ShowWindow newSession.statusWindow
210               newSession.connect
220           Else
230               If Not m_serverProfile Is Nothing Then
240                   newSession.serverProfile = m_serverProfile
250               Else
260                   newSession.serverHost = m_serverHost
270                   newSession.serverPort = m_serverPort
280                   newSession.primaryNickname = m_primaryNickname
290                   newSession.backupNickname = m_backupNickname
300                   newSession.realName = m_realName
310               End If
                  
320               m_client.ShowWindow newSession.statusWindow
330               newSession.connect
340           End If
350       Else
360           If LenB(server) <> 0 Then
370               Set profile = serverProfiles.findProfile(server)
                  
380               If Not profile Is Nothing Then
390                   Me.serverProfile = profile
400               Else
410                   Me.serverHost = server
420               End If
                  
430               If port <> 0 Then
440                   Me.serverPort = port
450               End If
                 
460               If LenB(password) <> 0 Then
470                   Me.serverPassword = password
480               End If
490           End If
              
500           connect
510       End If
End Sub

Private Sub splitServerPort(text As String, ByRef server As String, ByRef port As Long)
          Dim params() As String
          
10        params = Split(text, ":", 2)
          
20        If UBound(params) = 0 Then
30            server = params(0)
40            port = 0
50            Exit Sub
60        End If
          
70        server = params(0)
80        port = Val(params(1))
End Sub

Private Sub extractCommand(ByRef line As String, ByRef command As String)
          Dim slashes As Integer
          
10        Do While LenB(line) <> 0
20            Select Case left$(line, 1)
                  Case "/"
30                    slashes = slashes + 1
40                Case " "
50                    If LenB(command) <> 0 Then
60                        line = Mid$(line, 2)
70                        Exit Sub
80                    End If
90                Case Else
100                   command = command & left$(line, 1)
110           End Select
              
120           line = Mid$(line, 2)
130       Loop
End Sub

Private Function extractCommandParams(ByRef line As String, maxParams As Byte, ByRef params() As _
    String) As Byte
          Dim temp As String
          Dim paramCount As Byte

10        Do While LenB(line) <> 0
20            If left$(line, 1) = " " Then
30                If LenB(temp) <> 0 Then
40                    paramCount = paramCount + 1
50                    ReDim Preserve params(paramCount - 1)
60                    params(paramCount - 1) = temp
70                    temp = vbNullString
80                End If
90            Else
100               temp = temp & left$(line, 1)
110           End If
              
120           If paramCount + 1 >= maxParams Then
130               Do While (left$(line, 1) = " ")
140                   line = Mid$(line, 2)
150               Loop
                  
160               If LenB(line) = 0 Then
170                   extractCommandParams = paramCount
180                   Exit Function
190               End If
                  
200               paramCount = paramCount + 1
210               ReDim Preserve params(paramCount - 1)
220               params(paramCount - 1) = line
                  
230               extractCommandParams = paramCount
240               Exit Function
250           End If
              
260           line = Mid$(line, 2)
270       Loop
          
280       If LenB(temp) <> 0 Then
290           paramCount = paramCount + 1
300           ReDim Preserve params(paramCount - 1)
310           params(paramCount - 1) = temp
320       End If
          
330       extractCommandParams = paramCount
End Function

Private Sub initInputParsers()
10        addInputParser "JOIN", 1, 2, "commandJoin"
20        addInputParser "J", 1, 2, "commandJoin"
          
30        addInputParser "TOPIC", 0, 2, "commandTopic"
          
40        addInputParser "KICK", 1, 3, "commandKick"
50        addInputParser "K", 1, 3, "commandKick"
          
60        addInputParser "HOP", 0, 2, "commandHop"
70        addInputParser "REJOIN", 0, 2, "commandHop"
          
80        addInputParser "ECHO", 1, 1, "commandEcho"
          
90        addInputParser "QUOTE", 1, 1, "commandRaw"
100       addInputParser "RAW", 1, 1, "commandRaw"
          
110       addInputParser "PART", 0, 2, "commandPart"
          
120       addInputParser "OPTIONS", 0, 0, "commandOptions"
130       addInputParser "DEBUG", 0, 0, "commandDebug"
140       addInputParser "SERVER", 0, 4, "commandServer"
          
150       addInputParser "RECONNECT", 0, 0, "commandReconnect"
          
160       addInputParser "MSG", 2, 2, "commandMsg"
170       addInputParser "NOTICE", 2, 2, "commandNotice"
180       addInputParser "ME", 1, 1, "commandMe"
190       addInputParser "EMOTE", 1, 1, "commandMe"
200       addInputParser "ACTION", 1, 1, "commandMe"
210       addInputParser "QUERY", 1, 2, "commandQuery"
          
220       addInputParser "SLAP", 1, 1, "commandSlap"
230       addInputParser "HUGGLE", 1, 1, "commandHuggle"
          
240       addInputParser "CLEAR", 0, 0, "commandClear"
250       addInputParser "CLEARALL", 0, 0, "commandClearAll"
          
260       addInputParser "CTCP", 2, 3, "commandCtcp"
          
270       addInputParser "CLOSE", 0, 0, "commandClose"
280       addInputParser "CLOSEALL", 0, 0, "commandCloseAll"
          
290       addInputParser "LIST", 0, 2, "commandList"
300       addInputParser "QUIT", 0, 1, "commandQuit"
          
310       addInputParser "UNICODE", 0, 0, "commandUnicode"
          
320       addInputParser "IGNORE", 0, 3, "commandIgnore"
End Sub

Private Sub addInputParser(command As String, minParams As Byte, maxParams As Byte, func As String)
          Dim inputParser As New CInputParser
          
10        inputParser.init command, minParams, maxParams, func
          
20        m_inputParsers.Add inputParser, LCase$(command)
End Sub

Private Function findInputParser(command As String) As CInputParser
10        On Error Resume Next
20        Set findInputParser = m_inputParsers.item(LCase$(command))
End Function

Public Sub sendLine(line As String)
10        If Not m_debugWindowText Is Nothing Then
20            m_debugWindowText.addText "-> " & line
30        End If
          
40        m_socket.sendLine line
End Sub

Public Sub init()
10        Set m_reconnectTimer = m_client.getTimer
20        Set m_channelJoinTimer = m_client.getTimer
          
30        m_reconnectTimer.enabled = False
40        m_reconnectTimer.Interval = RECONNECT_DELAY
          
50        m_channelJoinTimer.enabled = False
60        m_channelJoinTimer.Interval = CHANNEL_JOIN_DELAY
          
70        Set m_socket = New CSwiftSocket
          
80        m_socket.init m_statusWindow.hwnd
End Sub

Private Sub Class_Initialize()
10        initParsers
20        initInputParsers
          
30        m_baseLogPath = g_userPath & LOG_DIR
End Sub

Public Sub deInit()
10        debugLog "Trying to deInit() CSession"

          Dim channel As CChannel
          Dim query As CQuery
          
20        debugLog "Destroying channels"
          
30        For Each channel In m_channels
40            destroyChannel channel
50        Next channel
          
60        debugLog "Destroying queries"
          
70        For Each query In m_queries
80            destroyQuery query
90        Next query
          
100       m_client.releaseTimer m_reconnectTimer
110       m_client.releaseTimer m_channelJoinTimer
          
120       m_socket.closeSocket
130       m_statusWindow.deInit
          
140       If Not m_debugWindow Is Nothing Then
150           closeGenericWindow m_debugWindow
160       End If
          
170       closeChannelList
          
180       Set m_statusWindow = Nothing
190       Set m_statusWindowText = Nothing
200       Set m_whoisText = Nothing
210       Set m_client = Nothing
End Sub

Private Sub m_socket_connected()
10        m_connected = True
20        Call login
End Sub

Private Sub m_socket_readable()
          Dim buf As String
          Dim result As Long
          
10        result = m_socket.sockRead(buf)
          
20        If result > 0 Then
30            parse buf
40        End If
End Sub

Public Sub channelTabbing(Source As swiftIrc.ctlTextInput, text As String, start As Long, Length As _
    Long)
          Dim matchLength As Long
          Dim count As Long
          Dim channel As CChannel
          
10        If m_lastChannelTabIndex <> 0 Then
20            If m_lastChannelTab = text Then
30                matchLength = Len(m_lastChannelTabMatch)
              
40                For count = m_lastChannelTabIndex + 1 To m_channels.count
50                    Set channel = m_channels.item(count)
                  
60                    If LCase(left$(channel.name, matchLength)) = m_lastChannelTabMatch Then
70                        m_lastChannelTab = channel.name
80                        m_lastChannelTabIndex = count
                          
90                        Source.replaceText start, Length, channel.name
100                       Exit Sub
110                   End If
120               Next count
                  
130               If m_lastChannelTabIndex > 1 Then
140                   For count = 1 To m_lastChannelTabIndex - 1
150                       Set channel = m_channels.item(count)
                  
160                       If LCase(left$(channel.name, matchLength)) = m_lastChannelTabMatch Then
170                           m_lastChannelTab = channel.name
180                           m_lastChannelTabIndex = count
                              
190                           Source.replaceText start, Length, channel.name
200                           Exit Sub
210                       End If
220                   Next count
230               End If
                  
240               Exit Sub
250           End If
260       End If
          
270       m_lastChannelTabMatch = LCase$(text)
280       matchLength = Len(text)
          
290       For count = 1 To m_channels.count
300           Set channel = m_channels.item(count)
                  
310           If LCase(left$(channel.name, matchLength)) = m_lastChannelTabMatch Then
320               m_lastChannelTab = channel.name
330               m_lastChannelTabIndex = count
                  
340               Source.replaceText start, Length, channel.name
350               Exit For
360           End If
370       Next count
End Sub

Public Sub registered()
10        If settings.enableFiltering Then
20            sendLine "MODE " & m_currentNickname & " +f"
30        End If
          
          Dim channels As String
          
40        If Not m_serverProfile Is Nothing Then
50            If m_serverProfile.enableAutoIdentify Then
60                sendLine "NICKSERV IDENTIFY " & m_serverProfile.nicknamePassword
70            End If
          
80            If m_serverProfile.enablePerform Then
90                runCommands m_serverProfile.perform
100           End If
110       End If
          
120       m_channelJoinTimer.enabled = True
End Sub

Private Sub runCommands(commands As String)
          Dim count As Long
          Dim line As String
          
10        For count = 1 To Len(commands)
20            If Mid$(commands, count, 1) = vbLf Or Mid$(commands, count, 1) = vbCr Then
30                If count < Len(commands) Then
40                    If Mid$(commands, count + 1, 1) = vbLf Or Mid$(commands, count + 1, 1) = vbCr Then
50                        count = count + 1
60                    End If
70                End If
                  
80                If left$(LTrim(line), 1) <> "/" Then
90                    line = "/" & line
100               End If
                  
110               line = Replace$(line, "$me", m_currentNickname, 1, -1, vbTextCompare)
120               textInput m_statusWindow, line
130               line = vbNullString
140           Else
150               line = line & Mid$(commands, count, 1)
160           End If
170       Next count
          
180       If LenB(line) <> 0 Then
190           If left$(LTrim(line), 1) <> "/" Then
200               line = "/" & line
210           End If
                  
220           line = Replace$(line, "$me", m_currentNickname, 1, -1, vbTextCompare)
230           textInput m_statusWindow, line
240       End If
End Sub

Private Sub joinChannels()
          Dim channels As String

10        If Not m_serverProfile Is Nothing Then
20            If m_serverProfile.enableAutoJoin Then
                  Dim ajChannel As CAutoJoinChannel
                  Dim count As Long
                  
30                For count = 1 To m_serverProfile.autoJoinChannelCount
40                    Set ajChannel = m_serverProfile.autoJoinChannel(count)
                  
50                    If findChannel(ajChannel.channel) Is Nothing Then
60                        If Len(channels) + Len(ajChannel.channel) + 8 _
                              >= 510 Then
70                            sendLine "JOIN " & channels
80                            channels = vbNullString
90                        End If
                      
100                       If LenB(ajChannel.key) <> 0 Then
110                           sendLine "JOIN " & channels
120                           sendLine "JOIN " & ajChannel.channel & " " & ajChannel.key
130                           channels = vbNullString
140                       Else
150                           channels = channels & ajChannel.channel & ","
160                       End If
170                   End If
180               Next count
190           End If
200       End If
          
          Dim channel As CChannel
          
210       For Each channel In m_channels
220           If Len(channels) + Len(channel.name) + 8 >= 510 Then
230               sendLine "JOIN " & channels
240               channels = vbNullString
250           End If
              
260           If LenB(channel.key) <> 0 Then
270               sendLine "JOIN " & channels
280               sendLine "JOIN " & channel.name & " " & channel.key
290               channels = vbNullString
300           Else
310               channels = channels & channel.name & ","
320           End If
330       Next channel
          
340       If LenB(channels) <> 0 Then
350           sendLine "JOIN " & channels
360       End If
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
          
10        params = Split(Trim(paramString), " ")
          
20        buffer = "MODE " & target & " "
          
30        modeOn = True
          
40        For count = 1 To Len(modes)
50            param = vbNullString
          
60            If modeCount >= m_maxModes Then
70                sendLine buffer & modeBuffer & " " & paramBuffer
                  
80                If modeOn = True Then
90                    modeBuffer = "+"
100               Else
110                   modeBuffer = "-"
120               End If
                  
130               paramBuffer = ""
140               modeCount = 0
150           End If
              
160           Select Case Mid$(modes, count, 1)
                  Case "+"
170                   modeOn = True
180               Case "-"
190                   modeOn = False
200               Case Else
210                   modeCount = modeCount + 1
                      
220                   If getPrefixModeIndex(Mid$(modes, count, 1)) <> 0 Then
230                       If paramIndex <= UBound(params) Then
240                           param = params(paramIndex)
250                       End If
260                   Else
270                       modeType = getChannelModeType(Mid$(modes, count, 1))
                          
280                       If modeType = cmtList Then
290                           If paramIndex <= UBound(params) Then
300                               param = params(paramIndex)
310                           End If
320                       ElseIf modeType = cmtParam Then
330                           If paramIndex <= UBound(params) Then
340                               param = params(paramIndex)
350                           End If
360                       ElseIf modeType = cmtSetOnly Then
370                           If modeOn Then
380                               If paramIndex <= UBound(params) Then
390                                   param = params(paramIndex)
400                               End If
410                           End If
420                       End If
430                   End If
                      
440                   If LenB(param) <> 0 Then
450                       paramIndex = paramIndex + 1
460                   End If
470           End Select
              
480           If Len(buffer) + (Len(modeBuffer) + 1) + Len(paramBuffer) + Len(param) + 2 > 500 Then
490               sendLine buffer & modeBuffer & " " & paramBuffer
                  
500               modeBuffer = IIf(modeOn = True, "+", "-") & Mid$(modes, count, 1)
510               paramBuffer = param
520               modeCount = 1
530           Else
540               modeBuffer = modeBuffer & Mid$(modes, count, 1)
                  
550               If LenB(param) <> 0 Then
560                   paramBuffer = paramBuffer & " " & param
570               End If
580           End If
590       Next count
          
600       sendLine buffer & modeBuffer & " " & paramBuffer
End Sub

Private Sub updateIal(origin As COrigin)
          Dim ialEntry As CIalEntry
          
10        Set ialEntry = findIal(origin.nickname)
          
20        If Not ialEntry Is Nothing Then
30            ialEntry.nick = origin.nickname
40            ialEntry.host = origin.hostname
50            ialEntry.ident = origin.username
60            Exit Sub
70        End If
          
80        Set ialEntry = New CIalEntry
          
90        ialEntry.nick = origin.nickname
100       ialEntry.ident = origin.username
110       ialEntry.host = origin.hostname
          
120       m_ial.Add ialEntry, LCase$(ialEntry.nick)
End Sub

Private Function findIal(nickname As String) As CIalEntry
10        On Error Resume Next
20        Set findIal = m_ial.item(LCase$(nickname))
End Function

Public Function getIal(nickname As String, ialType As eIalType) As String
          Dim ialEntry As CIalEntry
          
10        Set ialEntry = findIal(nickname)
          
20        If ialEntry Is Nothing Then
30            getIal = nickname & "!*@*"
40            Exit Function
50        End If
          
60        Select Case ialType
              Case ialAll
70                getIal = ialEntry.nick & "!" & ialEntry.ident & "@" & ialEntry.host
80            Case ialHost
90                getIal = "*!*@" & ialEntry.host
100           Case ialIdent
110               getIal = "*!*" & ialEntry.ident & "@*"
120           Case ialIdentHost
130               getIal = "*!*" & ialEntry.ident & "@" & ialEntry.host
140           Case Else
150               getIal = ialEntry.nick & "!" & ialEntry.ident & "@" & ialEntry.host
160       End Select
End Function

Public Function CCOpen() As Boolean
          ' This function will return true if the Channel Central
          ' is opened, false if not
          
          Dim frm As Form
          Dim result As Boolean
10        result = False
          
20        For Each frm In Forms
30            If frm.name = "frmChannelCentral" Then
40                result = True
50            End If
60        Next frm
          
70        CCOpen = result
End Function

