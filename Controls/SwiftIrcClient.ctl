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
Attribute VB_Exposed = False

Option Explicit

Private m_firstUseDialog As frmFirstUseDisclaimer

Private m_fontmanager As New CFontManager
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
    If Not m_firstUseDialog Is Nothing Then
        m_firstUseDialog.Show vbModeless, Me
        Exit Sub
    End If

    Set m_firstUseDialog = New frmFirstUseDisclaimer
    m_firstUseDialog.init Me
    m_firstUseDialog.Show vbModeless, Me
End Sub

Friend Sub agreementAccepted()
    Unload m_firstUseDialog
    Set m_firstUseDialog = Nothing
    
    Dim count As Long
    
    For count = 1 To m_sessions.count
        m_sessions.item(count).agreementAccepted
    Next count
End Sub

Friend Property Get fontManager() As CFontManager
    Set fontManager = m_fontmanager
End Property

Friend Property Get switchbar() As ctlSwitchbar
    Set switchbar = m_switchBar
End Property

Friend Property Get activeTextWindow() As ITextWindow
    If Not m_activeIWindow Is Nothing Then
        If TypeOf m_activeIWindow Is ITextWindow Then
            Set activeTextWindow = m_activeIWindow
        End If
    End If
End Property

Private Sub initStartPage()
    Set m_startPage = createNewWindow("swiftIrc.ctlWindowStart", "start")
    m_startPage.client = Me
    m_startPage.switchbartab = m_switchBar.addTab(Nothing, m_startPage, sboGeneric, "Start", Nothing)
    ShowWindow m_startPage
End Sub

Friend Function generateDefaultNickname() As String
    generateDefaultNickname = "SKUser" & CStr(Fix(1000 * Rnd))
End Function

Friend Function newSession() As CSession
    Dim statusWindow As ctlWindowStatus
    
    Set statusWindow = createNewWindow("swiftirc.ctlWindowStatus", "status")
    
    Dim session As New CSession
    
    session.client = Me
    session.statusWindow = statusWindow
    session.primaryNickname = generateDefaultNickname
    session.username = "swiftkituser"
    session.realName = "SwiftKitUser"
    
    statusWindow.session = session
    
    m_sessions.Add session
    Set newSession = session
    
    session.statusWindow.switchbartab = switchbar.addTab(Nothing, session.statusWindow, sboStatus, _
        "N/A", g_iconSBStatus)
    session.init
End Function

Friend Sub removeSession(session As CSession)
    m_switchBar.removeTab session.statusWindow.switchbartab
    session.statusWindow.switchbartab = Nothing
    destroyWindow session.statusWindow
    session.deInit
    
    Dim count As Long
    
    For count = 1 To m_sessions.count
        If m_sessions.item(count) Is session Then
            m_sessions.Remove count
            Exit For
        End If
    Next count
End Sub

Friend Function createNewWindow(progId As String, name As String) As IWindow
    Dim newControl As VBControlExtender
    
   On Error GoTo createNewWindow_Error

    Set newControl = Controls.Add(progId, name & m_nameCounter)
    
    Dim window As IWindow
    
    Set window = newControl.object
    window.realWindow = newControl
    
    m_nameCounter = m_nameCounter + 1
    
    If TypeOf window Is IFontUser Then
        Dim fontUser As IFontUser
        
        Set fontUser = window
        fontUser.fontManager = m_fontmanager
        fontUser.fontsUpdated
    End If
    
    If TypeOf window Is ITextWindow Then
        Dim textWindow As ITextWindow
        
        Set textWindow = window
        textWindow.eventManager = m_eventManager
        textWindow.update
    End If
    
    Set createNewWindow = window

   On Error GoTo 0
   Exit Function

createNewWindow_Error:
    handleError "createNewWindow", Err.Number, Err.Description, Erl, vbNullString
End Function

Friend Sub destroyWindow(window As IWindow)
    Controls.Remove window.realWindow.name
End Sub

Friend Sub ShowWindow(window As IWindow)
   On Error GoTo ShowWindow_Error

    If m_activeWindow Is window.realWindow Then
        Exit Sub
    End If
    
    Dim oldActiveWindow As VBControlExtender
    Set oldActiveWindow = m_activeWindow
    
    Set m_activeWindow = window.realWindow
    Set m_activeIWindow = window
    sizeActiveWindow
    m_activeWindow.visible = True
    
    If TypeOf window Is ITabWindow Then
        Dim tabWindow As ITabWindow
        
        Set tabWindow = window
        m_switchBar.selectTab tabWindow.getTab, False
    End If
    
    If TypeOf window Is ITextWindow Then
        Dim textWindow As ITextWindow
        
        Set textWindow = window
        
        textWindow.focusInput
    End If
    
    If Not oldActiveWindow Is Nothing Then
        oldActiveWindow.visible = False
    End If

   On Error GoTo 0
   Exit Sub

ShowWindow_Error:
    handleError "ShowWindow", Err.Number, Err.Description, Erl, vbNullString
End Sub

Friend Sub redrawSwitchbarTab(aTab As CTab)
    m_switchBar.redrawTab aTab
End Sub

Friend Function addSwitchbarTab(parentTab As CTab, window As IWindow, order As eSwitchbarOrder) As CTab
    Dim insertPos As Long
End Function

Friend Sub removeTab(aTab As CTab)
    m_switchBar.removeTab aTab
    aTab.window = Nothing
End Sub

Private Sub m_eventManager_themeUpdated()
    Dim aControl As control
    Dim textWindow As ITextWindow
    
    For Each aControl In Controls
        If TypeOf aControl Is VBControlExtender Then
            If TypeOf aControl.object Is ITextWindow Then
                Set textWindow = aControl.object
                textWindow.update
            End If
        End If
    Next aControl
End Sub

Private Sub m_startPage_profileConnect(profile As CServerProfile)
    Dim session As CSession
    
    Set session = newSession
    
    session.serverProfile = profile
    ShowWindow session.statusWindow
    session.connect
End Sub

Private Sub m_startPage_newSession()
    Dim session As CSession
    
    Set session = newSession
    ShowWindow session.statusWindow
End Sub

Private Sub m_switchBar_changeHeight(newHeight As Long)
    If newHeight > (UserControl.ScaleHeight / 2) Then
        newHeight = UserControl.ScaleHeight / 2
        
        Dim rows As Long
        
        rows = m_switchBar.getMaxRows(newHeight)
        m_switchBar.rows = rows
        Exit Sub
    End If

    If m_switchBar.position = sbpTop Then
        m_switchbarControl.Move 0, 0, UserControl.ScaleWidth, newHeight
    Else
        m_switchbarControl.Move 0, UserControl.ScaleHeight - newHeight, UserControl.ScaleWidth, _
            newHeight
    End If
    
    UserControl_Resize
End Sub

Private Sub m_switchBar_closeRequest(aTab As CTab)
    Dim session As CSession

    If TypeOf aTab.window Is ctlWindowStatus Then
        Dim statusWindow As ctlWindowStatus
        
        Set statusWindow = aTab.window
        removeSession statusWindow.session
    ElseIf TypeOf aTab.window Is ctlWindowChannel Then
        Dim channelWindow As ctlWindowChannel
        
        Set channelWindow = aTab.window
        Set session = channelWindow.session
        
        session.partChannel channelWindow.channel
    ElseIf TypeOf aTab.window Is ctlWindowQuery Then
        Dim queryWindow As ctlWindowQuery
        
        Set queryWindow = aTab.window
        Set session = queryWindow.session
        
        session.closeQuery queryWindow.query
    ElseIf TypeOf aTab.window Is ctlChannelList Then
        Dim channelList As ctlChannelList
        
        Set channelList = aTab.window
        Set session = channelList.session
        session.closeChannelList
    ElseIf TypeOf aTab.window Is ctlWindowGenericText Then
        Dim genericWindow As ctlWindowGenericText
        
        Set genericWindow = aTab.window
        Set session = genericWindow.session
        
        session.closeGenericWindow genericWindow
    End If
End Sub

Private Sub m_switchBar_moveRequest(x As Single, y As Single)
    Dim realY As Single
    
    realY = m_switchbarControl.top + y
    
    If realY >= UserControl.ScaleHeight - (UserControl.ScaleHeight / 3) Then
        If m_switchBar.position <> sbpBottom Then
            m_switchBar.position = sbpBottom
            UserControl_Resize
        End If
    ElseIf realY <= UserControl.ScaleHeight / 3 Then
        If m_switchBar.position <> sbpTop Then
            m_switchBar.position = sbpTop
            UserControl_Resize
        End If
    End If
End Sub

Private Sub m_switchBar_tabSelected(selectedTab As CTab)
    ShowWindow selectedTab.window
End Sub

Private Sub sizeActiveWindow()
    If Not m_activeWindow Is Nothing And Not m_switchbarControl Is Nothing Then
        If UserControl.ScaleHeight <= m_switchbarControl.height Then
            Exit Sub
        End If
    
        If m_switchBar.position = sbpTop Then
            m_activeWindow.Move 0, m_switchbarControl.height, UserControl.ScaleWidth, _
                UserControl.ScaleHeight - m_switchbarControl.height
        ElseIf m_switchBar.position = sbpBottom Then
            m_activeWindow.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - _
                m_switchbarControl.height
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    If Not m_switchBar Is Nothing Then
        If m_switchBar.position = sbpTop Then
            m_switchbarControl.Move 0, 0, UserControl.ScaleWidth, m_switchbarControl.height
        ElseIf m_switchBar.position = sbpBottom Then
            m_switchbarControl.Move 0, UserControl.ScaleHeight - m_switchbarControl.height, _
                UserControl.ScaleWidth, m_switchbarControl.height
        End If
    End If
    
    sizeActiveWindow
End Sub

Private Sub initFonts()
    If LenB(settings.fontName) = 0 Then
        changeFont getBestDefaultFont, 9, False, False
    Else
        If settings.fontSize = 0 Then
            changeFont settings.fontName, 9, settings.setting("fontBold", estBoolean), _
                settings.setting("fontItalic", estBoolean)
        Else
            changeFont settings.fontName, settings.fontSize, settings.setting("fontBold", _
                estBoolean), settings.setting("fontItalic", estBoolean)
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    If g_initialized Then
        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    End If
End Sub

Private Sub initUserStyles()
    Set prefixStyles = New cArrayList
    Dim imageOp As CImage
    Dim imageHalfOp As CImage
    Dim imageVoice As CImage
    
    Set imageOp = imageManager.addImage("op.bmp")
    Set imageHalfOp = imageManager.addImage("halfop.bmp")
    Set imageVoice = imageManager.addImage("voice.bmp")
    
    addPrefix "@", imageOp, settings.setting("nickColourOps", estNumber)
    addPrefix "%", imageHalfOp, settings.setting("nickColourHalfops", estNumber)
    addPrefix "+", imageVoice, settings.setting("nickColourVoices", estNumber)
    
    Set styleMe = New CUserStyle
    styleMe.init vbNullString, Nothing, settings.setting("nickColourMe", estNumber)
    
    Set styleMeOp = New CUserStyle
    styleMeOp.init "@", imageOp, settings.setting("nickColourMe", estNumber)
    
    Set styleMeHalfop = New CUserStyle
    styleMeHalfop.init "%", imageHalfOp, settings.setting("nickColourMe", estNumber)
    
    Set styleMeVoice = New CUserStyle
    styleMeVoice.init "+", imageVoice, settings.setting("nickColourMe", estNumber)
        
    Set styleNormal = New CUserStyle
    styleNormal.init vbNullString, Nothing, settings.setting("nickColourNormal", estNumber)
End Sub

Private Sub addPrefix(symbol As String, image As CImage, foreColour As Byte)
    Dim userStyle As New CUserStyle
    
    userStyle.init symbol, image, foreColour
    
    Dim prefix As New CPrefixStyle
    
    prefix.symbol = symbol
    prefix.style = userStyle
    
    prefixStyles.Add prefix
End Sub

Private Sub initImages()
    Set imageManager = New CImageManager
    imageManager.rootPath = g_AssetPath
    
    Set g_iconSBStatus = imageManager.addImage("status.bmp")
    Set g_iconSBChannel = imageManager.addImage("channel.bmp")
    Set g_iconSBQuery = imageManager.addImage("query.bmp")
    Set g_iconSBGeneric = imageManager.addImage("generic.bmp")
    Set g_iconSBList = imageManager.addImage("channel_list.bmp")
End Sub

Private Sub initEvents()
    If Not g_initialized Then
        Exit Sub
    End If

    Set m_eventManager = New CEventManager
    
    Dim classic As CTextTheme
    Dim modern As CTextTheme
    
    Set classic = New CTextTheme
    Set modern = New CTextTheme
    m_eventManager.defaultTheme = classic
    
    classic.addEvent "CONNECTING", "* Connecting to $0 ($1)...", eventColours.infoText, TVE_STANDARD
    classic.addEvent "CONNECTED", "* Connected to $0, logging in...", eventColours.infoText, _
        TVE_STANDARD
    classic.addEvent "DISCONNECTED", "* Disconnected, type /reconnect to reconnect", eventColours.infoText, TVE_STANDARD
    classic.addEvent "RECONNECTING_IN", "* Trying to reconnect in $0 second(s)", _
        eventColours.infoText, TVE_STANDARD
    
    classic.addEvent "IRC_ERROR", "* $0", eventColours.infoText, TVE_STANDARD
    
    classic.addEvent "KILLED", _
        "* You were disconnected from the server by operator $0 (Reason: $1)", eventColours.infoText, _
        TVE_STANDARD
    
    classic.addEvent "NICKNAME_IN_USE", "* The nickname $0 is already in use.", _
        eventColours.infoText, TVE_STANDARD
    classic.addEvent "NICKNAME_IN_USE_PREREG", _
        "* The nickname $0 is already in use.  Trying backup nickname instead...", _
        eventColours.infoText, TVE_STANDARD
    classic.addEvent "NICKNAME_IN_USE_PREREG2", _
        "* The nickname $0 is already in use.  Please enter a different nickname.", _
        eventColours.infoText, TVE_STANDARD
        
    classic.addEvent "NO_SUCH_NICK", "* No such nickname/channel ($0)", eventColours.infoText, _
        TVE_STANDARD
        
    classic.addEvent "WELCOME", "Logged onto $0 as $1", eventColours.infoText, TVE_STANDARD
    classic.addEvent "NUMERIC", "$0", eventColours.otherText, TVE_STANDARD
    
    classic.addEvent "PRIVMSG", "<$s> $0", eventColours.normalText, TVE_USERTEXT
    classic.addEvent "PRIVMSG_HIGHLIGHT", "<$s> $0", eventColours.highlightText, TVE_USERTEXT
    
    classic.addEvent "CHANNEL_PRIVMSG", "<$s> $0", eventColours.normalText, TVE_USERTEXT
    classic.addEvent "CHANNEL_PRIVMSG_HIGHLIGHT", "<$s> $0", eventColours.highlightText, TVE_USERTEXT
    classic.addEvent "WALLCHOP_PRIVMSG", "<$s:$0$1> $2", eventColours.normalText, TVE_USERTEXT
    
    classic.addEvent "ME_PRIVMSG", "<$s> $0", eventColours.myMessages, TVE_USERTEXT

    classic.addEvent "NOTICE", "-$0- $1", eventColours.noticeText, TVE_USERTEXT Or TVE_SEPERATE_BOTH
    classic.addEvent "CHANNEL_NOTICE", "-$s:$0- $1", eventColours.noticeText, TVE_USERTEXT
    classic.addEvent "WALLCHOP_NOTICE", "-$s:$0$1- $2", eventColours.noticeText, TVE_USERTEXT
    
    classic.addEvent "EMOTE", "* $0 $1", eventColours.emotes, TVE_USERTEXT
    classic.addEvent "EMOTE_HIGHLIGHT", "* $0 $1", eventColours.highlightText, TVE_USERTEXT
    classic.addEvent "CHANNEL_EMOTE", "* $s $0", eventColours.emotes, TVE_USERTEXT
    classic.addEvent "CHANNEL_EMOTE_HIGHLIGHT", "* $s $0", eventColours.highlightText, TVE_USERTEXT
    classic.addEvent "WALLCHOP_EMOTE", "* $s:$0$1 $2", eventColours.emotes, TVE_USERTEXT
    
    classic.addEvent "CTCP_RECEIVED", "[$0 $1] $2", eventColours.ctcpText, TVE_USERTEXT Or TVE_TIMESTAMP
    classic.addEvent "CTCP_REPLY_RECEIVED", "[$0 $1 reply]: $2", eventColours.ctcpText, TVE_USERTEXT Or TVE_TIMESTAMP
    
    classic.addEvent "ME_JOIN", "* Now talking in $0", eventColours.channelJoin, TVE_STANDARD
    classic.addEvent "ME_PART", "* You have left $0", eventColours.channelPart, TVE_STANDARD
    classic.addEvent "ME_REJOINING", "* Attempting to rejoin $0...", eventColours.infoText, _
        TVE_STANDARD
    classic.addEvent "ME_REJOINED", "* Rejoined channel $0", eventColours.channelJoin, TVE_STANDARD
    classic.addEvent "ME_KICKED", "* You were kicked from $0 by $1 ($2$o)", _
        eventColours.channelKick, TVE_USERTEXT
        
    classic.addEvent "ME_REJOIN_DELAY", "* Attempting to rejoin $0 in $1...", eventColours.infoText, TVE_STANDARD
    
    classic.addEvent "USER_JOIN", "* $0 ($2@$3) has joined $1", eventColours.channelJoin, _
        TVE_STANDARD
    classic.addEvent "USER_PART", "* $0 ($2@$3) has left $1", eventColours.channelPart, TVE_STANDARD
    classic.addEvent "USER_PART_REASON", "* $0 ($3@$4) has left $1 ($2$o)", _
        eventColours.channelPart, TVE_USERTEXT
    classic.addEvent "USER_QUIT", "* $s has quit IRC", eventColours.quit, TVE_STANDARD
    classic.addEvent "USER_QUIT_EX", "* $s ($0@$1) has quit IRC", eventColours.quit, TVE_STANDARD
    classic.addEvent "USER_QUIT_REASON", "* $s has quit IRC ($0$o)", eventColours.quit, TVE_USERTEXT
    classic.addEvent "USER_QUIT_REASON_EX", "* $s ($0@$1) has quit IRC ($2$o)", eventColours.quit, _
        TVE_USERTEXT
    classic.addEvent "USER_NICK_CHANGE", "* $0 is now known as $1", eventColours.nickChanges, _
        TVE_STANDARD
    classic.addEvent "USER_KICKED", "* $2 was kicked by $0 ($3$o)", eventColours.channelKick, _
        TVE_USERTEXT
    
    classic.addEvent "CHANNEL_MODE_CHANGE", "* $0 set mode: $2 $3", eventColours.modeChange, _
        TVE_STANDARD
    'classic.addEvent "CHANNEL_MODE_OP", "$0 made $1 a channel operator", eventColours.modeChange, _
        TVE_STANDARD
    'classic.addEvent "CHANNEL_MODE_DEOP", "$0 took $1's operator status", eventColours.modeChange, _
        TVE_STANDARD
    
    classic.addEvent "CHANNEL_TOPICIS", "* Topic is '$0$o'", eventColours.topicChange, TVE_USERTEXT
    classic.addEvent "CHANNEL_TOPICWHOTIME", "* Set by $0 on $1", eventColours.topicChange, _
        TVE_STANDARD
    classic.addEvent "CHANNEL_TOPICCHANGE", "* $0 changed the topic to '$1$o'", _
        eventColours.topicChange, TVE_USERTEXT
    
    classic.addEvent "ERROR_CONNECT", "Error connecting to $0: $1", eventColours.infoText, _
        TVE_STANDARD
    classic.addEvent "ERROR_DISCONNECT", "Disconnected with error: $0", eventColours.infoText, _
        TVE_STANDARD
    
    classic.addEvent "WHOIS_USER", "$0 is $1@$2 * $3", eventColours.whoisText, TVE_VISIBLE Or _
        TVE_INDENTWRAP Or TVE_SEPERATE_TOP Or TVE_SEPERATE_EXPLICIT
    classic.addEvent "WHOIS_CHANNELS", "$0 is on $1", eventColours.whoisText, TVE_VISIBLE Or _
        TVE_INDENTWRAP
    classic.addEvent "WHOIS_SERVER", "$0 using server $1 ($2)", eventColours.whoisText, TVE_VISIBLE _
        Or TVE_INDENTWRAP
    classic.addEvent "WHOIS_OPERATOR", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or _
        TVE_INDENTWRAP
    classic.addEvent "WHOIS_IDLE", "$0 has been idle $1, signed on $2", eventColours.whoisText, _
        TVE_VISIBLE Or TVE_INDENTWRAP
    classic.addEvent "WHOIS_REGNICK", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP
    classic.addEvent "WHOIS_GENERIC", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP
    classic.addEvent "WHOIS_END", "$0 $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP Or _
        TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "AWAY", "$0 is away: $1", eventColours.whoisText, TVE_VISIBLE Or TVE_INDENTWRAP
    
    classic.addEvent "ERR_BANNEDFROMCHAN", "* Cannot join $0, you are banned from the channel", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_INVITEONLYCHAN", "* Cannot join $0, channel is invite only", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_NEEDREGGEDNICK", _
        "* Cannot join $0, you need to be logged into a registered nickname", eventColours.infoText, _
        TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_BADCHANNELKEY", "* Cannot join $0, wrong channel key/password was given", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_CHANNELISFULL", "* Cannot join $0, channel is full", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_TOOMANYJOINS", _
        "* Cannot join $0, you are trying to rejoin the channel too quickly", eventColours.infoText, _
        TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_TOOMANYCHANNELS", "* Cannot join $0, you are already in too many channels", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_SECUREONLYCHAN", _
        "* Cannot join $0, you need to be using an encrypted (SSL) connection", eventColours.infoText, _
        TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "ERR_NOPRIVILEGES", _
        "* Permission denied, you do not have the correct IRC operator privileges to execute this action", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    
    classic.addEvent "CMD_INSUFFICIENT_PARAMS", "* /$0: Insufficient parameters", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    classic.addEvent "CMD_INCOMPATIBLE_WINDOW", "* /$0: Can not use command $0 in this window", _
        eventColours.infoText, TVE_VISIBLE Or TVE_SEPERATE_BOTH
    
    classic.addEvent "CMD_PRIVMSG_SENT", "-> *$0* $1", eventColours.otherText, TVE_USERTEXT Or TVE_SEPERATE_BOTH
    classic.addEvent "CMD_NOTICE_SENT", "-> -$0- $1", eventColours.otherText, TVE_USERTEXT Or TVE_SEPERATE_BOTH
    classic.addEvent "CMD_RAW_SENT", "-> Server: $0", eventColours.otherText, TVE_VISIBLE Or _
        TVE_SEPERATE_BOTH
        
    classic.addEvent "CMD_CTCP_SENT", "-> [$0] $1", eventColours.ctcpText, TVE_VISIBLE Or TVE_SEPERATE_BOTH Or TVE_TIMESTAMP
      
    classic.addEvent "CHANNEL_CTCP", "[$s:$0 $1] $2", eventColours.ctcpText, TVE_USERTEXT
    classic.addEvent "WALLCHOP_CTCP", "[$s:$0$1 $2] $3", eventColours.ctcpText, TVE_USERTEXT
    
    classic.addEvent "ME_MODE_CHANGE", "$0 sets mode: $1", eventColours.modeChange, _
        TVE_STANDARD Or TVE_SEPERATE_BOTH
        
    classic.addEvent "IGNORE_LIST_START", "- Ignore list -", eventColours.infoText, TVE_STANDARD
    classic.addEvent "IGNORE_LIST_ENTRY", "* $0 ($1)", eventColours.infoText, TVE_STANDARD
    classic.addEvent "IGNORE_LIST_END", " - End of ignore list -", eventColours.infoText, TVE_STANDARD
    
    classic.addEvent "IGNORE_ADDED", "* Added $0 ($1) to ignore list", eventColours.infoText, TVE_STANDARD
    classic.addEvent "IGNORE_UPDATED", "* Updated ignore list item: $0 ($1)", eventColours.infoText, TVE_STANDARD
    classic.addEvent "IGNORE_REMOVED", "* Removed $0 from ignore list", eventColours.infoText, TVE_STANDARD
    classic.addEvent "IGNORE_REMOVE_NOTFOUND", "* $0 not found on ignore list", eventColours.infoText, TVE_STANDARD
    classic.addEvent "IGNORE_LIST_CLEARED", "* Cleared ignore list", eventColours.infoText, TVE_STANDARD
    
    classic.addEvent "IGNORE_INVALID_FLAGS", "* Invalid ignore flags ($0)", eventColours.infoText, TVE_STANDARD
    classic.addEvent "IGNORE_INVALID_COMMAND", "* Invalid ignore operation ($0)", eventColours.infoText, TVE_STANDARD
        
    ' 10/nov/2011
    classic.addEvent "WHO_LIST", "$0 $1!$2@$3 $4 $5", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "WHO_END_OF_LIST", "* $0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "WHOWAS_HOST", "$0 was $1!$2 * $3", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_EXPLICIT
    classic.addEvent "WHOWAS_UNKNOWN", "$0: $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_EXPLICIT
    classic.addEvent "WHOWAS_END", "* $0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "SILENCE_LIST", "$0", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "END_OF_SILENCE_LIST", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "ERR_CHGNICK_MODEN", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    classic.addEvent "ERR_CHGNICK_MODEB", "$1 ($0)", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "HELPOP_TITLE", "$0", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "HELPOP_TEXT", "$0", eventColours.otherText, TVE_VISIBLE
    
    classic.addEvent "MARKED_AWAY", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    classic.addEvent "NO_LONGER_MARKED_AWAY", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "INVITE_LIST", "$0", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "END_OF_INVITE_LIST", "$0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "INVITE_USER", "You've invited $0 to $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "BAN_LIST", "$0 $1 set by $2 on $3", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "END_OF_BAN_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

    classic.addEvent "EX_LIST", "$0 $1 set by $2 on $3", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "END_OF_EX_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

    classic.addEvent "INVEX_LIST", "$0 $1 set by $2 on $3", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "END_OF_INVEX_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

    classic.addEvent "A_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "END_OF_A_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "Q_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE
    classic.addEvent "END_OF_Q_LIST", "$0 $1", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "SILENCE_MODIFY", "SILENCE $0", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT
    
    classic.addEvent "INVITATION_RECEIVED", "You were invited to $0 by $1 ($2@$3)", eventColours.otherText, TVE_VISIBLE Or TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM Or TVE_SEPERATE_EXPLICIT

    m_eventManager.loadTheme classic
End Sub

Friend Sub changeFont(fontName As String, fontSize As Integer, fontBold As Boolean, fontItalic As Boolean)
    m_fontmanager.changeFont UserControl.hdc, fontName, fontSize, fontBold, fontItalic

    Dim aControl As control
    Dim fontUser As IFontUser
    
    For Each aControl In Controls
        If TypeOf aControl Is VBControlExtender Then
            If TypeOf aControl.object Is IFontUser Then
                Set fontUser = aControl.object
                fontUser.fontsUpdated
            End If
        End If
    Next aControl
    
    settings.fontName = fontName
    settings.fontSize = fontSize
    settings.setting("fontBold", estBoolean) = fontBold
    settings.setting("fontItalic", estBoolean) = fontItalic
    settings.saveSettings
End Sub

Public Function getTimer() As Object
    Set getTimer = Controls.Add("VB.Timer", "timer" & m_timerCounter)
    m_timerCounter = m_timerCounter + 1
End Function

Public Sub releaseTimer(timer As Object)
    Controls.Remove timer.name
End Sub

Public Property Get appHandlesUrls() As Boolean
    appHandlesUrls = m_appHandlesUrls
End Property

Public Property Let appHandlesUrls(newValue As Boolean)
    m_appHandlesUrls = newValue
End Property

Public Property Get colourWindow() As Long
    colourWindow = colourManager.getColour(SWIFTCOLOUR_WINDOW)
End Property

Public Property Let colourWindow(newValue As Long)
    colourManager.setColour SWIFTCOLOUR_WINDOW, newValue
    coloursUpdated
End Property

Public Property Get colourControlBack() As Long
    colourControlBack = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
    coloursUpdated
End Property

Public Property Let colourControlBack(newValue As Long)
    colourManager.setColour SWIFTCOLOUR_CONTROLBACK, newValue
    coloursUpdated
End Property

Public Property Get colourControlFore() As Long
    colourControlFore = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
End Property

Public Property Let colourControlFore(newValue As Long)
    colourManager.setColour SWIFTCOLOUR_CONTROLFORE, newValue
    coloursUpdated
End Property

Public Property Get colourControlForeOver() As Long
    colourControlForeOver = colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
End Property

Public Property Let colourControlForeOver(newValue As Long)
    colourManager.setColour SWIFTCOLOUR_CONTROLFOREOVER, newValue
    coloursUpdated
End Property

Public Property Get colourControlBorder() As Long
    colourControlBorder = colourManager.getColour(SWIFTCOLOUR_CONTROLBORDER)
End Property

Public Property Let colourControlBorder(newValue As Long)
    colourManager.setColour SWIFTCOLOUR_CONTROLBORDER, newValue
    coloursUpdated
End Property
    
Public Property Get colourFrameBack() As Long
    colourFrameBack = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
End Property

Public Property Let colourFrameBack(newValue As Long)
    colourManager.setColour SWIFTCOLOUR_FRAMEBACK, newValue
    colourManager.setColour SWIFTCOLOUR_FRAMEBORDER, newValue
    coloursUpdated
End Property

Public Property Get assetPath() As String
    assetPath = g_AssetPath
End Property

Public Property Let assetPath(newValue As String)
    g_AssetPath = newValue
End Property

Public Property Get userPath() As String
    userPath = g_userPath
End Property

Public Property Let userPath(newValue As String)
    g_userPath = newValue
End Property

Public Property Get debugEx() As Boolean
    debugEx = g_debugModeEx
End Property

Public Property Let debugEx(newValue As Boolean)
    g_debugModeEx = newValue
End Property

Public Sub init()
    On Error GoTo init_error:

    g_clientCount = g_clientCount + 1

    If Not g_initialized Then
        If LenB(Dir(g_AssetPath & "hand.cur", vbNormal)) <> 0 Then
            Set g_handCursor = LoadPicture(assetPath & "hand.cur")
        End If
        
        Set settings = New CSettings
        Set colourManager = New CColourManager
    
        settings.loadSettings
        
        initUIFonts UserControl.hdc
        
        g_initialized = True
    
        Randomize
        
        Set eventColours = New CEventColours
        
        Set textManager = New CTextManager
        Set serverProfiles = New CServerProfileManager
        Set colourThemes = New CColourThemeManager
        Set highlights = New CHighlightManager
        Set ignoreManager = New CIgnoreManager
        
        colourThemes.loadThemes
        serverProfiles.loadProfiles
        textManager.loadText
        
        eventColours.loadTheme colourThemes.currentTheme
        
        ignoreManager.loadIgnoreList
        
        g_textViewBack = colourThemes.currentTheme.backgroundColour
        g_textViewFore = eventColours.normalText.colour
        
        g_textInputBack = colourThemes.currentTheme.backgroundColour
        g_textInputFore = eventColours.normalText.colour
        
        g_nicklistBack = colourThemes.currentTheme.backgroundColour
        g_nicklistFore = eventColours.normalText.colour
        
        g_channelListBack = colourThemes.currentTheme.backgroundColour
        g_channelListFore = eventColours.normalText.colour
        
        highlights.load
        
        initImages
        initUserStyles
        
    End If
    
    initFonts
    
    initEvents
    initLogging
    
    Dim window As IWindow

    Set m_switchBar = createNewWindow("swiftIrc.ctlSwitchbar", "switchbar")
    Set m_switchbarControl = getRealWindow(m_switchBar)

    m_switchbarControl.visible = True
    m_switchbarControl.Move 0, 0, UserControl.ScaleWidth, m_switchBar.getRequiredHeight

    initStartPage
    
    UserControl_Resize
    
    If StrComp(settings.setting("switchbarPosition", estString), "bottom", vbTextCompare) = 0 Then
        m_switchBar.position = sbpBottom
    Else
        m_switchBar.position = sbpTop
    End If
    
    m_switchBar.rows = settings.setting("switchbarRows", estNumber)
    
    registerForOptionsUpdates Me
    
    Exit Sub
    
init_error:
    Resume Next
End Sub

Public Sub initLogging()
    On Error Resume Next

    If Dir(g_userPath & "logs", vbDirectory) = vbNullString Then
        MkDir g_userPath & "logs"
    End If
End Sub

Public Sub deInit()
    If isOptionsFormParent(Me) Then
        closeOptionsDialog
    End If
    
    settings.saveSettings

    Dim count As Long
    
    For count = m_sessions.count To 1 Step -1
        removeSession m_sessions.item(count)
    Next count
    
    Dim a As CSession
    
    Set m_activeIWindow = Nothing
    Set m_activeWindow = Nothing

    g_clientCount = g_clientCount - 1
    
    If g_clientCount = 0 Then
        g_initialized = False
    End If
    
    unregisterForOptionsUpdates Me
End Sub

Public Function getVersion() As String
    getVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Friend Function findSwiftIRCSession() As CSession
    Dim count As Long
    
    For count = 1 To m_sessions.count
        If m_sessions.item(count).networkName = "SwiftIRC" Then
            Set findSwiftIRCSession = m_sessions.item(count)
            Exit Function
        End If
    Next count
End Function

Friend Sub coloursUpdated()
    updateColours Controls
End Sub

Friend Sub refreshFontSettings()
    Me.changeFont settings.fontName, settings.fontSize, settings.setting("fontBold", estBoolean), settings.setting("fontItalic", estBoolean)
End Sub

Friend Sub refreshSwitchbarSettings()
    If StrComp(settings.setting("switchbarPosition", estString), "Top", vbTextCompare) = 0 Then
        m_switchBar.position = sbpTop
    ElseIf StrComp(settings.setting("switchbarPosition", estString), "Bottom", vbTextCompare) = 0 Then
        m_switchBar.position = sbpBottom
    End If
    
    m_switchBar.rows = settings.setting("switchbarRows", estNumber)
End Sub

Friend Sub visitUrl(url As String, Optional unSafe As Boolean = True)
    If unSafe Then
        Dim result As Long
        
        result = MsgBox("The web address:" & vbCrLf & vbCrLf & url & vbCrLf & vbCrLf & "may not be safe, are you sure you wish to visit it?", vbQuestion Or vbYesNo, "Visit URL")
    
        If result <> vbYes Then
            Exit Sub
        End If
    End If

    If m_appHandlesUrls Then
        RaiseEvent visitUrl(url)
        Exit Sub
    End If
    
    launchDefaultBrowser url
End Sub

Friend Sub showIgnoreList(session As CSession)
    Dim ignoreList As New frmIgnoreList
    
    ignoreList.session = session
    ignoreList.Show vbModal, Me
    
    Unload ignoreList
End Sub

