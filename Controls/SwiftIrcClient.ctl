VERSION 5.00
Begin VB.UserControl SwiftIrcClient 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   Begin VB.Menu mnuTab 
      Caption         =   "mnuTab"
      Begin VB.Menu mnuTabSwitch 
         Caption         =   "Switch to"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "SwiftIrcClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private m_firstUseDialog As frmFirstUseDisclaimer

Private m_fontManager As New CFontManager
Private m_activeIWindow As IWindow
Private m_activeWindow As VBControlExtender
Private m_nameCounter As Long
Private m_timerCounter As Long
Private m_sockCounter As Long

Private m_sessions As New cArrayList

Private WithEvents m_eventManager As CEventManager
Attribute m_eventManager.VB_VarHelpID = -1

Private WithEvents m_switchBar As ctlSwitchbar
Attribute m_switchBar.VB_VarHelpID = -1
Private m_switchbarControl As VBControlExtender

Private WithEvents m_startPage As ctlWindowStart
Attribute m_startPage.VB_VarHelpID = -1

Private m_appHandlesUrls As Boolean

Public Event visitUrl(url As String)

Friend Sub showUserAgreement()
10        If Not m_firstUseDialog Is Nothing Then
20            m_firstUseDialog.Show vbModeless, Me
30            Exit Sub
40        End If

50        Set m_firstUseDialog = New frmFirstUseDisclaimer
60        m_firstUseDialog.init Me
70        m_firstUseDialog.Show vbModeless, Me
End Sub

Friend Sub agreementAccepted()
10        Unload m_firstUseDialog
20        Set m_firstUseDialog = Nothing
          
          Dim count As Long
          
30        For count = 1 To m_sessions.count
40            m_sessions.item(count).agreementAccepted
50        Next count
End Sub

Friend Property Get fontManager() As CFontManager
10        Set fontManager = m_fontManager
End Property

Friend Property Get switchbar() As ctlSwitchbar
10        Set switchbar = m_switchBar
End Property

Friend Property Get activeTextWindow() As ITextWindow
10        If Not m_activeIWindow Is Nothing Then
20            If TypeOf m_activeIWindow Is ITextWindow Then
30                Set activeTextWindow = m_activeIWindow
40            End If
50        End If
End Property

Private Sub initStartPage()
10        Set m_startPage = createNewWindow("swiftIrc.ctlWindowStart", "start")
20        m_startPage.client = Me
30        m_startPage.switchbartab = m_switchBar.addTab(Nothing, m_startPage, sboGeneric, "Start", Nothing)
40        ShowWindow m_startPage
End Sub

Friend Function generateDefaultNickname() As String
10        generateDefaultNickname = "SKUser" & CStr(Fix(1000 * Rnd))
End Function

Friend Function newSession() As CSession
          Dim statusWindow As ctlWindowStatus
          
10        Set statusWindow = createNewWindow("swiftirc.ctlWindowStatus", "status")
          
          Dim session As New CSession
          
20        session.client = Me
30        session.statusWindow = statusWindow
40        session.primaryNickname = generateDefaultNickname
50        session.username = "swiftkituser"
60        session.realName = "SwiftKitUser"
          
70        statusWindow.session = session
          
80        m_sessions.Add session
90        Set newSession = session
          
100       session.statusWindow.switchbartab = switchbar.addTab(Nothing, session.statusWindow, sboStatus, _
              "N/A", g_iconSBStatus)
110       session.init
End Function

Friend Sub removeSession(session As CSession)
10        m_switchBar.removeTab session.statusWindow.switchbartab
20        session.statusWindow.switchbartab = Nothing
30        destroyWindow session.statusWindow
40        session.deInit
          
          Dim count As Long
          
50        For count = 1 To m_sessions.count
60            If m_sessions.item(count) Is session Then
70                m_sessions.Remove count
80                Exit For
90            End If
100       Next count
End Sub

Friend Function createNewWindow(progId As String, name As String) As IWindow
          Dim newControl As VBControlExtender
          
10       On Error GoTo createNewWindow_Error

20        Set newControl = Controls.Add(progId, name & m_nameCounter)
          
          Dim window As IWindow
          
30        Set window = newControl.object
40        window.realWindow = newControl
          
50        m_nameCounter = m_nameCounter + 1
          
60        If TypeOf window Is IFontUser Then
              Dim fontUser As IFontUser
              
70            Set fontUser = window
80            fontUser.fontManager = m_fontManager
90            fontUser.fontsUpdated
100       End If
          
110       If TypeOf window Is ITextWindow Then
              Dim textWindow As ITextWindow
              
120           Set textWindow = window
130           textWindow.eventManager = m_eventManager
140           textWindow.update
150       End If
          
160       Set createNewWindow = window

170      On Error GoTo 0
180      Exit Function

createNewWindow_Error:
190       handleError "createNewWindow", Err.Number, Err.Description, Erl, vbNullString
End Function

Friend Sub destroyWindow(window As IWindow)
10        Controls.Remove window.realWindow.name
End Sub

Friend Sub ShowWindow(window As IWindow)
10       On Error GoTo ShowWindow_Error

20        If m_activeWindow Is window.realWindow Then
30            Exit Sub
40        End If
          
          Dim oldActiveWindow As VBControlExtender
50        Set oldActiveWindow = m_activeWindow
          
60        Set m_activeWindow = window.realWindow
70        Set m_activeIWindow = window
80        sizeActiveWindow
90        m_activeWindow.visible = True
          
100       If TypeOf window Is ITabWindow Then
              Dim tabWindow As ITabWindow
              
110           Set tabWindow = window
120           m_switchBar.selectTab tabWindow.getTab, False
130       End If
          
140       If TypeOf window Is ITextWindow Then
              Dim textWindow As ITextWindow
              
150           Set textWindow = window
              
160           textWindow.focusInput
170       End If
          
180       If Not oldActiveWindow Is Nothing Then
190           oldActiveWindow.visible = False
200       End If

210      On Error GoTo 0
220      Exit Sub

ShowWindow_Error:
230       handleError "ShowWindow", Err.Number, Err.Description, Erl, vbNullString
End Sub

Friend Sub redrawSwitchbarTab(aTab As CTab)
10        m_switchBar.redrawTab aTab
End Sub

Friend Function addSwitchbarTab(parentTab As CTab, window As IWindow, order As eSwitchbarOrder) As CTab
    Dim insertPos As Long
End Function

Friend Sub removeTab(aTab As CTab)
10        m_switchBar.removeTab aTab
20        aTab.window = Nothing
End Sub

Private Sub m_eventManager_themeUpdated()
          Dim aControl As control
          Dim textWindow As ITextWindow
          
10        For Each aControl In Controls
20            If TypeOf aControl Is VBControlExtender Then
30                If TypeOf aControl.object Is ITextWindow Then
40                    Set textWindow = aControl.object
50                    textWindow.update
60                End If
70            End If
80        Next aControl
End Sub

Private Sub m_startPage_profileConnect(profile As CServerProfile)
          Dim session As CSession
          
10        Set session = newSession
          
20        session.serverProfile = profile
30        ShowWindow session.statusWindow
40        session.connect
End Sub

Private Sub m_startPage_newSession()
          Dim session As CSession
          
10        Set session = newSession
20        ShowWindow session.statusWindow
End Sub

Private Sub m_switchBar_changeHeight(newHeight As Long)
10        If newHeight > (UserControl.ScaleHeight / 2) Then
20            newHeight = UserControl.ScaleHeight / 2
              
              Dim rows As Long
              
30            rows = m_switchBar.getMaxRows(newHeight)
40            m_switchBar.rows = rows
50            Exit Sub
60        End If

70        If m_switchBar.position = sbpTop Then
80            m_switchbarControl.Move 0, 0, UserControl.ScaleWidth, newHeight
90        Else
100           m_switchbarControl.Move 0, UserControl.ScaleHeight - newHeight, UserControl.ScaleWidth, _
                  newHeight
110       End If
          
120       UserControl_Resize
End Sub

Private Sub m_switchBar_closeRequest(aTab As CTab)
          Dim session As CSession

10        If TypeOf aTab.window Is ctlWindowStatus Then
              Dim statusWindow As ctlWindowStatus
              
20            Set statusWindow = aTab.window
30            removeSession statusWindow.session
40        ElseIf TypeOf aTab.window Is ctlWindowChannel Then
              Dim channelWindow As ctlWindowChannel
              
50            Set channelWindow = aTab.window
60            Set session = channelWindow.session
              
70            session.partChannel channelWindow.channel
80        ElseIf TypeOf aTab.window Is ctlWindowQuery Then
              Dim queryWindow As ctlWindowQuery
              
90            Set queryWindow = aTab.window
100           Set session = queryWindow.session
              
110           session.closeQuery queryWindow.query
120       ElseIf TypeOf aTab.window Is ctlChannelList Then
              Dim channelList As ctlChannelList
              
130           Set channelList = aTab.window
140           Set session = channelList.session
150           session.closeChannelList
160       ElseIf TypeOf aTab.window Is ctlWindowGenericText Then
              Dim genericWindow As ctlWindowGenericText
              
170           Set genericWindow = aTab.window
180           Set session = genericWindow.session
              
190           session.closeGenericWindow genericWindow
200       End If
End Sub

Private Sub m_switchBar_moveRequest(x As Single, y As Single)
          Dim realY As Single
          
10        realY = m_switchbarControl.top + y
          
20        If realY >= UserControl.ScaleHeight - (UserControl.ScaleHeight / 3) Then
30            If m_switchBar.position <> sbpBottom Then
40                m_switchBar.position = sbpBottom
50                UserControl_Resize
60            End If
70        ElseIf realY <= UserControl.ScaleHeight / 3 Then
80            If m_switchBar.position <> sbpTop Then
90                m_switchBar.position = sbpTop
100               UserControl_Resize
110           End If
120       End If
End Sub

Private Sub m_switchBar_tabSelected(selectedTab As CTab)
10        ShowWindow selectedTab.window
End Sub

Private Sub sizeActiveWindow()
10        If Not m_activeWindow Is Nothing And Not m_switchbarControl Is Nothing Then
20            If UserControl.ScaleHeight <= m_switchbarControl.height Then
30                Exit Sub
40            End If
          
50            If m_switchBar.position = sbpTop Then
60                m_activeWindow.Move 0, m_switchbarControl.height, UserControl.ScaleWidth, _
                      UserControl.ScaleHeight - m_switchbarControl.height
70            ElseIf m_switchBar.position = sbpBottom Then
80                m_activeWindow.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - _
                      m_switchbarControl.height
90            End If
100       End If
End Sub

Private Sub UserControl_Resize()
10        If Not m_switchBar Is Nothing Then
20            If m_switchBar.position = sbpTop Then
30                m_switchbarControl.Move 0, 0, UserControl.ScaleWidth, m_switchbarControl.height
40            ElseIf m_switchBar.position = sbpBottom Then
50                m_switchbarControl.Move 0, UserControl.ScaleHeight - m_switchbarControl.height, _
                      UserControl.ScaleWidth, m_switchbarControl.height
60            End If
70        End If
          
80        sizeActiveWindow
End Sub

Private Sub initFonts()
10        If LenB(settings.fontName) = 0 Then
20            changeFont getBestDefaultFont, 9, False, False
30        Else
40            If settings.fontSize = 0 Then
50                changeFont settings.fontName, 9, settings.setting("fontBold", estBoolean), _
                      settings.setting("fontItalic", estBoolean)
60            Else
70                changeFont settings.fontName, settings.fontSize, settings.setting("fontBold", _
                      estBoolean), settings.setting("fontItalic", estBoolean)
80            End If
90        End If
End Sub

Private Sub UserControl_Initialize()
10        If g_initialized Then
20            UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
30        End If
End Sub

Private Sub initUserStyles()
10        Set prefixStyles = New cArrayList
          Dim imageOp As CImage
          Dim imageHalfOp As CImage
          Dim imageVoice As CImage
          
20        Set imageOp = imageManager.addImage("op.bmp")
30        Set imageHalfOp = imageManager.addImage("halfop.bmp")
40        Set imageVoice = imageManager.addImage("voice.bmp")
          
50        addPrefix "@", imageOp, settings.setting("nickColourOps", estNumber)
60        addPrefix "%", imageHalfOp, settings.setting("nickColourHalfops", estNumber)
70        addPrefix "+", imageVoice, settings.setting("nickColourVoices", estNumber)
          
80        Set styleMe = New CUserStyle
90        styleMe.init vbNullString, Nothing, settings.setting("nickColourMe", estNumber)
          
100       Set styleMeOp = New CUserStyle
110       styleMeOp.init "@", imageOp, settings.setting("nickColourMe", estNumber)
          
120       Set styleMeHalfop = New CUserStyle
130       styleMeHalfop.init "%", imageHalfOp, settings.setting("nickColourMe", estNumber)
          
140       Set styleMeVoice = New CUserStyle
150       styleMeVoice.init "+", imageVoice, settings.setting("nickColourMe", estNumber)
              
160       Set styleNormal = New CUserStyle
170       styleNormal.init vbNullString, Nothing, settings.setting("nickColourNormal", estNumber)
End Sub

Private Sub addPrefix(symbol As String, image As CImage, foreColour As Byte)
          Dim userStyle As New CUserStyle
          
10        userStyle.init symbol, image, foreColour
          
          Dim prefix As New CPrefixStyle
          
20        prefix.symbol = symbol
30        prefix.style = userStyle
          
40        prefixStyles.Add prefix
End Sub

Private Sub initImages()
10        Set imageManager = New CImageManager
20        imageManager.rootPath = g_AssetPath
          
30        Set g_iconSBStatus = imageManager.addImage("status.bmp")
40        Set g_iconSBChannel = imageManager.addImage("channel.bmp")
50        Set g_iconSBQuery = imageManager.addImage("query.bmp")
60        Set g_iconSBGeneric = imageManager.addImage("generic.bmp")
70        Set g_iconSBList = imageManager.addImage("channel_list.bmp")
End Sub

Private Sub initEvents()
10        If Not g_initialized Then
20            Exit Sub
30        End If

40        Set m_eventManager = New CEventManager
          
          Dim classic As CTextTheme
          Dim modern As CTextTheme
          
50        Set classic = New CTextTheme
60        Set modern = New CTextTheme
70        m_eventManager.defaultTheme = classic
          
80        classic.addEvent "CONNECTING", "* Connecting to $0 ($1)...", eventColours.infoText, TVE_STANDARD
90        classic.addEvent "CONNECTED", "* Connected to $0, logging in...", eventColours.infoText, _
              TVE_STANDARD
100       classic.addEvent "DISCONNECTED", "* Disconnected, type /reconnect to reconnect", eventColours.infoText, TVE_STANDARD
110       classic.addEvent "RECONNECTING_IN", "* Trying to reconnect in $0 second(s)", _
              eventColours.infoText, TVE_STANDARD
          
120       classic.addEvent "IRC_ERROR", "* $0", eventColours.infoText, TVE_STANDARD
          
130       classic.addEvent "KILLED", _
              "* You were disconnected from the server by operator $0 (Reason: $1)", eventColours.infoText, _
              TVE_STANDARD
          
140       classic.addEvent "NICKNAME_IN_USE", "* The nickname $0 is already in use.", _
              eventColours.infoText, TVE_STANDARD
150       classic.addEvent "NICKNAME_IN_USE_PREREG", _
              "* The nickname $0 is already in use.  Trying backup nickname instead...", _
              eventColours.infoText, TVE_STANDARD
160       classic.addEvent "NICKNAME_IN_USE_PREREG2", _
              "* The nickname $0 is already in use.  Please enter a different nickname.", _
              eventColours.infoText, TVE_STANDARD
              
170       classic.addEvent "NO_SUCH_NICK", "* No such nickname/channel ($0)", eventColours.infoText, _
              TVE_STANDARD
              
180       classic.addEvent "WELCOME", "Logged onto $0 as $1", eventColours.infoText, TVE_STANDARD
190       classic.addEvent "NUMERIC", "$0", eventColours.otherText, TVE_STANDARD
          
200       classic.addEvent "PRIVMSG", "<$s> $0", eventColours.normalText, TVE_USERTEXT
210       classic.addEvent "PRIVMSG_HIGHLIGHT", "<$s> $0", eventColours.highlightText, TVE_USERTEXT
          
220       classic.addEvent "CHANNEL_PRIVMSG", "<$s> $0", eventColours.normalText, TVE_USERTEXT
230       classic.addEvent "CHANNEL_PRIVMSG_HIGHLIGHT", "<$s> $0", eventColours.highlightText, TVE_USERTEXT
240       classic.addEvent "WALLCHOP_PRIVMSG", "<$s:$0$1> $2", eventColours.normalText, TVE_USERTEXT
          
250       classic.addEvent "ME_PRIVMSG", "<$s> $0", eventColours.myMessages, TVE_USERTEXT

260       classic.addEvent "NOTICE", "-$0- $1", eventColours.noticeText, TVE_USERTEXT Or TVE_SEPERATE_BOTH
270       classic.addEvent "CHANNEL_NOTICE", "-$s:$0- $1", eventColours.noticeText, TVE_USERTEXT
280       classic.addEvent "WALLCHOP_NOTICE", "-$s:$0$1- $2", eventColours.noticeText, TVE_USERTEXT
          
290       classic.addEvent "EMOTE", "* $0 $1", eventColours.emotes, TVE_USERTEXT
300       classic.addEvent "EMOTE_HIGHLIGHT", "* $0 $1", eventColours.highlightText, TVE_USERTEXT
310       classic.addEvent "CHANNEL_EMOTE", "* $s $0", eventColours.emotes, TVE_USERTEXT
320       classic.addEvent "CHANNEL_EMOTE_HIGHLIGHT", "* $s $0", eventColours.highlightText, TVE_USERTEXT
330       classic.addEvent "WALLCHOP_EMOTE", "* $s:$0$1 $2", eventColours.emotes, TVE_USERTEXT
          
340       classic.addEvent "CTCP_RECEIVED", "[$0 $1] $2", eventColours.ctcpText, TVE_USERTEXT Or TVE_TIMESTAMP
350       classic.addEvent "CTCP_REPLY_RECEIVED", "[$0 $1 reply]: $2", eventColours.ctcpText, TVE_USERTEXT Or TVE_TIMESTAMP
          
360       classic.addEvent "ME_JOIN", "* Now talking in $0", eventColours.channelJoin, TVE_STANDARD
370       classic.addEvent "ME_PART", "* You have left $0", eventColours.channelPart, TVE_STANDARD
380       classic.addEvent "ME_REJOINING", "* Attempting to rejoin $0...", eventColours.infoText, _
              TVE_STANDARD
390       classic.addEvent "ME_REJOINED", "* Rejoined channel $0", eventColours.channelJoin, TVE_STANDARD
400       classic.addEvent "ME_KICKED", "* You were kicked from $0 by $1 ($2$o)", _
              eventColours.channelKick, TVE_USERTEXT
              
410       classic.addEvent "ME_REJOIN_DELAY", "* Attempting to rejoin $0 in $1...", eventColours.infoText, TVE_STANDARD
          
420       classic.addEvent "USER_JOIN", "* $0 ($2@$3) has joined $1", eventColours.channelJoin, _
              TVE_STANDARD
430       classic.addEvent "USER_PART", "* $0 ($2@$3) has left $1", eventColours.channelPart, TVE_STANDARD
440       classic.addEvent "USER_PART_REASON", "* $0 ($3@$4) has left $1 ($2$o)", _
              eventColours.channelPart, TVE_USERTEXT
450       classic.addEvent "USER_QUIT", "* $s has quit IRC", eventColours.quit, TVE_STANDARD
460       classic.addEvent "USER_QUIT_EX", "* $s ($0@$1) has quit IRC", eventColours.quit, TVE_STANDARD
470       classic.addEvent "USER_QUIT_REASON", "* $s has quit IRC ($0$o)", eventColours.quit, TVE_USERTEXT
480       classic.addEvent "USER_QUIT_REASON_EX", "* $s ($0@$1) has quit IRC ($2$o)", eventColours.quit, _
              TVE_USERTEXT
490       classic.addEvent "USER_NICK_CHANGE", "* $0 is now known as $1", eventColours.nickChanges, _
              TVE_STANDARD
500       classic.addEvent "USER_KICKED", "* $2 was kicked by $0 ($3$o)", eventColours.channelKick, _
              TVE_USERTEXT
          
510       classic.addEvent "CHANNEL_MODE_CHANGE", "* $0 set mode: $2 $3", eventColours.modeChange, _
              TVE_STANDARD
          'classic.addEvent "CHANNEL_MODE_OP", "$0 made $1 a channel operator", eventColours.modeChange, _
              TVE_STANDARD
          'classic.addEvent "CHANNEL_MODE_DEOP", "$0 took $1's operator status", eventColours.modeChange, _
              TVE_STANDARD
          
520       classic.addEvent "CHANNEL_TOPICIS", "* Topic is '$0$o'", eventColours.topicChange, TVE_USERTEXT
530       classic.addEvent "CHANNEL_TOPICWHOTIME", "* Set by $0 on $1", eventColours.topicChange, _
              TVE_STANDARD
540       classic.addEvent "CHANNEL_TOPICCHANGE", "* $0 changed the topic to '$1$o'", _
              eventColours.topicChange, TVE_USERTEXT
          
550       classic.addEvent "ERROR_CONNECT", "Error connecting to $0: $1", eventColours.infoText, _
              TVE_STANDARD
560       classic.addEvent "ERROR_DISCONNECT", "Disconnected with error: $0", eventColours.infoText, _
              TVE_STANDARD
          
570       classic.addEvent "WHOIS_USER", "$0 is $1@$2 * $3", eventColours.whoisText, TVE_VISIBLE Or _
              TVE_INDENTWRAP Or TVE_SEPERATE_TOP Or TVE_SEPERATE_EXPLICIT
580       classic.addEvent "WHOIS_CHANNELS", "$0 is on $1", eventColours.whoisText, TVE_VISIBLE Or _
              TVE_INDENTWRAP
590       classic.addEvent "WHOIS_SERVER", "$0 using server $1 ($2)", eventColours.whoisText, TVE_VISIBLE _
              Or TVE_INDENTWRAP
600       classic.addEvent "WHOIS_OPERATOR", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or _
              TVE_INDENTWRAP
610       classic.addEvent "WHOIS_IDLE", "$0 has been idle $1, signed on $2", eventColours.whoisText, _
              TVE_VISIBLE Or TVE_INDENTWRAP
620       classic.addEvent "WHOIS_REGNICK", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP
630       classic.addEvent "WHOIS_GENERIC", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP
640       classic.addEvent "WHOIS_END", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP Or _
              TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
650       classic.addEvent "AWAY", "$0 is away: $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP
          
660       classic.addEvent "ERR_BANNEDFROMCHAN", "* Cannot join $0, you are banned from the channel", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
670       classic.addEvent "ERR_INVITEONLYCHAN", "* Cannot join $0, channel is invite only", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
680       classic.addEvent "ERR_NEEDREGGEDNICK", _
              "* Cannot join $0, you need to be logged into a registered nickname", eventColours.infoText, _
              TVE_VISIBLE Or TVE_SEPERATE_BOTH
690       classic.addEvent "ERR_BADCHANNELKEY", "* Cannot join $0, wrong channel key/password was given", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
700       classic.addEvent "ERR_CHANNELISFULL", "* Cannot join $0, channel is full", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
710       classic.addEvent "ERR_TOOMANYJOINS", _
              "* Cannot join $0, you are trying to rejoin the channel too quickly", eventColours.infoText, _
              TVE_VISIBLE Or TVE_SEPERATE_BOTH
720       classic.addEvent "ERR_TOOMANYCHANNELS", "* Cannot join $0, you are already in too many channels", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
730       classic.addEvent "ERR_SECUREONLYCHAN", _
              "* Cannot join $0, you need to be using an encrypted (SSL) connection", eventColours.infoText, _
              TVE_VISIBLE Or TVE_SEPERATE_BOTH
740       classic.addEvent "ERR_NOPRIVILEGES", _
              "* Permission denied, you do not have the correct IRC operator privileges to execute this action", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
          
750       classic.addEvent "CMD_INSUFFICIENT_PARAMS", "* /$0: Insufficient parameters", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
760       classic.addEvent "CMD_INCOMPATIBLE_WINDOW", "* /$0: Can not use command $0 in this window", _
              eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
          
770       classic.addEvent "CMD_PRIVMSG_SENT", "-> *$0* $1", eventColours.otherText, TVE_USERTEXT Or TVE_SEPERATE_BOTH
780       classic.addEvent "CMD_NOTICE_SENT", "-> -$0- $1", eventColours.otherText, TVE_USERTEXT Or TVE_SEPERATE_BOTH
790       classic.addEvent "CMD_RAW_SENT", "-> Server: $0", eventColours.otherText, TVE_VISIBLE Or _
              TVE_SEPERATE_BOTH
              
800       classic.addEvent "CMD_CTCP_SENT", "-> [$0] $1", eventColours.ctcpText, TVE_VISIBLE Or TVE_SEPERATE_BOTH Or TVE_TIMESTAMP
            
810       classic.addEvent "CHANNEL_CTCP", "[$s:$0 $1] $2", eventColours.ctcpText, TVE_USERTEXT
820       classic.addEvent "WALLCHOP_CTCP", "[$s:$0$1 $2] $3", eventColours.ctcpText, TVE_USERTEXT
          
830       classic.addEvent "ME_MODE_CHANGE", "$0 sets mode: $1", eventColours.modeChange, _
              TVE_STANDARD Or TVE_SEPERATE_BOTH
              
840       classic.addEvent "IGNORE_LIST_START", "- Ignore list -", eventColours.infoText, TVE_STANDARD
850       classic.addEvent "IGNORE_LIST_ENTRY", "* $0 ($1)", eventColours.infoText, TVE_STANDARD
860       classic.addEvent "IGNORE_LIST_END", " - End of ignore list -", eventColours.infoText, TVE_STANDARD
          
870       classic.addEvent "IGNORE_ADDED", "* Added $0 ($1) to ignore list", eventColours.infoText, TVE_STANDARD
880       classic.addEvent "IGNORE_UPDATED", "* Updated ignore list item: $0 ($1)", eventColours.infoText, TVE_STANDARD
890       classic.addEvent "IGNORE_REMOVED", "* Removed $0 from ignore list", eventColours.infoText, TVE_STANDARD
900       classic.addEvent "IGNORE_REMOVE_NOTFOUND", "* $0 not found on ignore list", eventColours.infoText, TVE_STANDARD
910       classic.addEvent "IGNORE_LIST_CLEARED", "* Cleared ignore list", eventColours.infoText, TVE_STANDARD
          
920       classic.addEvent "IGNORE_INVALID_FLAGS", "* Invalid ignore flags ($0)", eventColours.infoText, TVE_STANDARD
930       classic.addEvent "IGNORE_INVALID_COMMAND", "* Invalid ignore operation ($0)", eventColours.infoText, TVE_STANDARD
              
          ' 10/nov/2011
940       classic.addEvent "WHO_LIST", "$0 $1!$2@$3 $4 $5", eventColours.otherText, TVE_VISIBLE
950       classic.addEvent "WHO_END_OF_LIST", "* $0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
960       classic.addEvent "WHOWAS_HOST", "$0 was $1!$2 * $3", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_EXPLICIT
970       classic.addEvent "WHOWAS_UNKNOWN", "$0: $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_EXPLICIT
980       classic.addEvent "WHOWAS_END", "* $0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
990       classic.addEvent "SILENCE_LIST", "$0", eventColours.otherText, TVE_VISIBLE
1000      classic.addEvent "END_OF_SILENCE_LIST", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1010      classic.addEvent "ERR_CHGNICK_MODEN", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
1020      classic.addEvent "ERR_CHGNICK_MODEB", "$1 ($0)", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1030      classic.addEvent "HELPOP_TITLE", "$0", eventColours.otherText, TVE_VISIBLE
1040      classic.addEvent "HELPOP_TEXT", "$0", eventColours.otherText, TVE_VISIBLE
          
1050      classic.addEvent "MARKED_AWAY", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
1060      classic.addEvent "NO_LONGER_MARKED_AWAY", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1070      classic.addEvent "INVITE_LIST", "$0", eventColours.otherText, TVE_VISIBLE
1080      classic.addEvent "END_OF_INVITE_LIST", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1090      classic.addEvent "INVITE_USER", "You've invited $0 to $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1100      classic.addEvent "BAN_LIST", "$0 $1 set by $2 on $3", eventColours.otherText, TVE_VISIBLE
1110      classic.addEvent "END_OF_BAN_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

1120      classic.addEvent "EX_LIST", "$0 $1 set by $2 on $3", eventColours.otherText, TVE_VISIBLE
1130      classic.addEvent "END_OF_EX_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

1140      classic.addEvent "INVEX_LIST", "$0 $1 set by $2 on $3", eventColours.otherText, TVE_VISIBLE
1150      classic.addEvent "END_OF_INVEX_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

1160      classic.addEvent "A_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE
1170      classic.addEvent "END_OF_A_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1180      classic.addEvent "Q_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE
1190      classic.addEvent "END_OF_Q_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1200      classic.addEvent "SILENCE_MODIFY", "SILENCE $0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
          
1210      classic.addEvent "INVITATION_RECEIVED", "You were invited to $0 by $1 ($2@$3)", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

1220      m_eventManager.loadTheme classic
End Sub

Friend Sub changeFont(fontName As String, fontSize As Integer, fontBold As Boolean, fontItalic As Boolean)
10        m_fontManager.changeFont UserControl.hdc, fontName, fontSize, fontBold, fontItalic

          Dim aControl As control
          Dim fontUser As IFontUser
          
20        For Each aControl In Controls
30            If TypeOf aControl Is VBControlExtender Then
40                If TypeOf aControl.object Is IFontUser Then
50                    Set fontUser = aControl.object
60                    fontUser.fontsUpdated
70                End If
80            End If
90        Next aControl
          
100       settings.fontName = fontName
110       settings.fontSize = fontSize
120       settings.setting("fontBold", estBoolean) = fontBold
130       settings.setting("fontItalic", estBoolean) = fontItalic
140       settings.saveSettings
End Sub

Private Sub UserControl_Terminate()
10        debugLog "swiftIrcClient terminating"
End Sub

Public Function getTimer() As Object
10        Set getTimer = Controls.Add("VB.Timer", "timer" & m_timerCounter)
20        m_timerCounter = m_timerCounter + 1
End Function

Public Sub releaseTimer(timer As Object)
10        Controls.Remove timer.name
End Sub

Public Property Get appHandlesUrls() As Boolean
10        appHandlesUrls = m_appHandlesUrls
End Property

Public Property Let appHandlesUrls(newValue As Boolean)
10        m_appHandlesUrls = newValue
End Property

Public Property Get colourWindow() As Long
10        colourWindow = colourManager.getColour(SWIFTCOLOUR_WINDOW)
End Property

Public Property Let colourWindow(newValue As Long)
10        colourManager.setColour SWIFTCOLOUR_WINDOW, newValue
20        coloursUpdated
End Property

Public Property Get colourControlBack() As Long
10        colourControlBack = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
20        coloursUpdated
End Property

Public Property Let colourControlBack(newValue As Long)
10        colourManager.setColour SWIFTCOLOUR_CONTROLBACK, newValue
20        coloursUpdated
End Property

Public Property Get colourControlFore() As Long
10        colourControlFore = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
End Property

Public Property Let colourControlFore(newValue As Long)
10        colourManager.setColour SWIFTCOLOUR_CONTROLFORE, newValue
20        coloursUpdated
End Property

Public Property Get colourControlForeOver() As Long
10        colourControlForeOver = colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
End Property

Public Property Let colourControlForeOver(newValue As Long)
10        colourManager.setColour SWIFTCOLOUR_CONTROLFOREOVER, newValue
20        coloursUpdated
End Property

Public Property Get colourControlBorder() As Long
10        colourControlBorder = colourManager.getColour(SWIFTCOLOUR_CONTROLBORDER)
End Property

Public Property Let colourControlBorder(newValue As Long)
10        colourManager.setColour SWIFTCOLOUR_CONTROLBORDER, newValue
20        coloursUpdated
End Property
    
Public Property Get colourFrameBack() As Long
10        colourFrameBack = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
End Property

Public Property Let colourFrameBack(newValue As Long)
10        colourManager.setColour SWIFTCOLOUR_FRAMEBACK, newValue
20        colourManager.setColour SWIFTCOLOUR_FRAMEBORDER, newValue
30        coloursUpdated
End Property

Public Property Get assetPath() As String
10        assetPath = g_AssetPath
End Property

Public Property Let assetPath(newValue As String)
10        g_AssetPath = newValue
End Property

Public Property Get userPath() As String
10        userPath = g_userPath
End Property

Public Property Let userPath(newValue As String)
10        g_userPath = newValue
End Property

Public Property Get debugEx() As Boolean
10        debugEx = g_debugModeEx
End Property

Public Property Let debugEx(newValue As Boolean)
10        g_debugModeEx = newValue
End Property

Public Sub init()
    On Error GoTo init_error:

10        debugLogEx "Init() called"

20        g_clientCount = g_clientCount + 1

30        If Not g_initialized Then
40            debugLogEx "First time initialization"
          
50            If LenB(Dir(g_AssetPath & "hand.cur", vbNormal)) <> 0 Then
60                debugLogEx "Found hand cursor, loading"
70                Set g_handCursor = LoadPicture(assetPath & "hand.cur")
80            End If
              
90            debugLogEx "Init settings object"
100           Set settings = New CSettings
110           Set colourManager = New CColourManager
          
120           debugLogEx "Loading settings..."
130           settings.loadSettings
              
140           debugLogEx "Init UI fonts"
150           initUIFonts UserControl.hdc
              
160           g_initialized = True
          
170           Randomize
              
180           debugLogEx "Init event colours object"
190           Set eventColours = New CEventColours
              
200           debugLogEx "Init text manager object"
210           Set textManager = New CTextManager
220           debugLogEx "Init server profiles object"
230           Set serverProfiles = New CServerProfileManager
240           debugLogEx "Init colour theme manager object"
250           Set colourThemes = New CColourThemeManager
260           debugLogEx "Init highlight manager object"
270           Set highlights = New CHighlightManager
280           Set ignoreManager = New CIgnoreManager
              
290           debugLogEx "Load colour themes"
300           colourThemes.loadThemes g_userPath & "swiftirc_themes.xml"
310           debugLogEx "Load server profiles"
320           serverProfiles.loadProfiles g_userPath & "swiftirc_servers.xml"
330           debugLogEx "Load text"
340           textManager.loadText g_AssetPath & "swiftirc_text.xml"
              
350           debugLogEx "Load current theme."
360           eventColours.loadTheme colourThemes.currentTheme
              
370           ignoreManager.loadIgnoreList g_userPath & "swiftirc_ignore_list.xml"
              
380           debugLogEx "Set custom control colours."
390           g_textViewBack = colourThemes.currentTheme.backgroundColour
400           g_textViewFore = eventColours.normalText.colour
              
410           g_textInputBack = colourThemes.currentTheme.backgroundColour
420           g_textInputFore = eventColours.normalText.colour
              
430           g_nicklistBack = colourThemes.currentTheme.backgroundColour
440           g_nicklistFore = eventColours.normalText.colour
              
450           g_channelListBack = colourThemes.currentTheme.backgroundColour
460           g_channelListFore = eventColours.normalText.colour
              
470           debugLogEx "Load highlights"
480           highlights.load
              
490           debugLogEx "Init images"
500           initImages
510           debugLogEx "Init user styles"
520           initUserStyles
              
530           debugLogEx "App wide init done."
540       End If
          
550       debugLogEx "Init client fonts"
560       initFonts
          
570       debugLogEx "Init events"
580       initEvents
590       debugLogEx "Init logging"
600       initLogging
          
          Dim window As IWindow

610       debugLogEx "Create switchbar"
620       Set m_switchBar = createNewWindow("swiftIrc.ctlSwitchbar", "switchbar")
630       Set m_switchbarControl = getRealWindow(m_switchBar)

640       m_switchbarControl.visible = True
650       m_switchbarControl.Move 0, 0, UserControl.ScaleWidth, m_switchBar.getRequiredHeight

660       debugLogEx "Init start page"
670       initStartPage
          
680       debugLogEx "Force resize."
690       UserControl_Resize
          
700       debugLogEx "Set switchbar position"
710       If StrComp(settings.setting("switchbarPosition", estString), "bottom", vbTextCompare) = 0 Then
720           m_switchBar.position = sbpBottom
730       Else
740           m_switchBar.position = sbpTop
750       End If
          
760       debugLogEx "Set switchbar rows"
770       m_switchBar.rows = settings.setting("switchbarRows", estNumber)
          
780       Exit Sub
          
init_error:
790       debugLogEx "Error: " & Err.Number & " (" & Err.Description & ") on line " & Erl
800       Resume Next
End Sub

Public Sub initLogging()
10        On Error Resume Next

20        If Dir(g_userPath & "logs", vbDirectory) = vbNullString Then
30            MkDir g_userPath & "logs"
40        End If
End Sub

Public Sub deInit()
10        settings.saveSettings

20        debugLog "Entering deInit()"

          Dim count As Long
          
30        For count = m_sessions.count To 1 Step -1
40            debugLog "Trying to deInit and remove session named " & m_sessions.item(count).networkName
50            removeSession m_sessions.item(count)
60        Next count
          
          Dim a As CSession
          
70        Set m_activeIWindow = Nothing
80        Set m_activeWindow = Nothing
          
90        debugLog "Done clearing sessions.  m_sessions.count = " & m_sessions.count
100       debugLog CStr(m_switchBar.tabCount)
          
110       g_clientCount = g_clientCount - 1
          
120       If g_clientCount = 0 Then
130           g_initialized = False
140       End If
End Sub

Public Function getVersion() As String
10        getVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Friend Function findSwiftIRCSession() As CSession
          Dim count As Long
          
10        For count = 1 To m_sessions.count
20            If m_sessions.item(count).networkName = "SwiftIRC" Then
30                Set findSwiftIRCSession = m_sessions.item(count)
40                Exit Function
50            End If
60        Next count
End Function

Friend Sub coloursUpdated()
10        updateColours Controls
End Sub

Friend Sub visitUrl(url As String, Optional unSafe As Boolean = True)
10        If unSafe Then
              Dim result As Long
              
20            result = MsgBox("The web address:" & vbCrLf & vbCrLf & url & vbCrLf & vbCrLf & "may not be safe, are you sure you wish to visit it?", vbQuestion Or vbYesNo, "Visit URL")
          
30            If result <> vbYes Then
40                Exit Sub
50            End If
60        End If

70        If m_appHandlesUrls Then
80            RaiseEvent visitUrl(url)
90            Exit Sub
100       End If
          
110       launchDefaultBrowser url
End Sub

Friend Sub showIgnoreList(session As CSession)
          Dim ignoreList As New frmIgnoreList
          
10        ignoreList.session = session
20        ignoreList.Show vbModal, Me
          
30        Unload ignoreList
End Sub

