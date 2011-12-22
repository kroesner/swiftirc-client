VERSION 5.00
Begin VB.UserControl ctlWindowChannel 
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   244
   Begin VB.Menu menuNicklist 
      Caption         =   "menuNicklist"
      Visible         =   0   'False
      Begin VB.Menu menuQuery 
         Caption         =   "Query"
      End
      Begin VB.Menu menuWhois 
         Caption         =   "Whois"
      End
      Begin VB.Menu menuControl 
         Caption         =   "Control"
         Begin VB.Menu menuOp 
            Caption         =   "Op"
         End
         Begin VB.Menu menuDeop 
            Caption         =   "Deop"
         End
         Begin VB.Menu menuHalfop 
            Caption         =   "Halfop"
         End
         Begin VB.Menu menuDehalfop 
            Caption         =   "Dehalfop"
         End
         Begin VB.Menu menuVoice 
            Caption         =   "Voice"
         End
         Begin VB.Menu menuDevoice 
            Caption         =   "Devoice"
         End
         Begin VB.Menu menuKick 
            Caption         =   "Kick"
         End
         Begin VB.Menu menuKickReason 
            Caption         =   "Kick (with reason)"
         End
         Begin VB.Menu menuBan 
            Caption         =   "Ban"
         End
         Begin VB.Menu menuKickBan 
            Caption         =   "Ban and kick"
         End
         Begin VB.Menu menuKickBanReason 
            Caption         =   "Ban and kick (with reason)"
         End
      End
      Begin VB.Menu mnuCtcp 
         Caption         =   "Ctcp"
         Begin VB.Menu mnuCtcpPing 
            Caption         =   "Ping"
         End
         Begin VB.Menu mnuCtcpVersion 
            Caption         =   "Version"
         End
         Begin VB.Menu mnuCtcpTime 
            Caption         =   "Time"
         End
      End
      Begin VB.Menu menuActions 
         Caption         =   "Actions"
         Begin VB.Menu menuSlap 
            Caption         =   "Slap"
         End
         Begin VB.Menu menuHuggle 
            Caption         =   "Huggle"
         End
      End
      Begin VB.Menu menuIgnore 
         Caption         =   "Ignore"
      End
   End
End
Attribute VB_Name = "ctlWindowChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements ITextWindow
Implements ITabWindow
Implements IFontUser
Implements IColourUser

Private m_nicklistWidth As Long
Private m_textInputHeight As Long

Private WithEvents m_nicklist As ctlNickList
Attribute m_nicklist.VB_VarHelpID = -1
Private WithEvents m_textView As ctlTextView
Attribute m_textView.VB_VarHelpID = -1
Private WithEvents m_textInput As ctlTextInput
Attribute m_textInput.VB_VarHelpID = -1

Private m_nameCounter As Long

Private m_fontManager As CFontManager

Private m_children As New cArrayList

Private m_tab As CTab 'our switchbar tab

'Nickname completer
Private m_lastTabIndex As Long
Private m_lastTab As String
Private m_lastTabMatch As String

Private m_list As New cArrayList
Private m_realWindow As VBControlExtender
Private m_session As CSession
Private m_channel As CChannel

Private m_selectedNick As CNicklistItem

Private Sub m_textInput_pageDown()
    m_textView.pageDown
End Sub

Private Sub m_textInput_pageUp()
    m_textView.pageUp
End Sub


Private Property Get ITextWindow_textview() As ctlTextView
    Set ITextWindow_textview = m_textView
End Property

Private Sub ITextWindow_clear()
    m_textView.clear
End Sub

Private Sub m_textView_clickedUrl(url As String)
    If left$(url, 1) = "#" Then
        If InStr(url, ",0") = 0 Then
            m_session.joinChannel url
        End If
    Else
        m_session.client.visitUrl url, True
    End If
End Sub

Private Sub m_textView_noLongerNeedFocus()
    If getRealWindow(m_textInput).visible Then
        getRealWindow(m_textInput).setFocus
    End If
End Sub

Private Sub m_textInput_mouseWheel(delta As Long)
    m_textView.processMouseWheel delta
End Sub

Private Sub ITextWindow_focusInput()
    If getRealWindow(m_textInput).visible Then
        getRealWindow(m_textInput).setFocus
    End If
End Sub

Private Sub IColourUser_coloursUpdated()
    updateColours Controls
End Sub

Public Sub init(session As CSession, channel As CChannel)
    Set m_session = session
    Set m_channel = channel
    m_nicklist.session = session
End Sub

Public Sub deInit()
    Set m_session = Nothing
    Set m_channel = Nothing
    m_nicklist.session = Nothing
    Set m_realWindow = Nothing
End Sub

Public Property Get switchbartab() As CTab
    Set switchbartab = m_tab
End Property

Public Property Let switchbartab(newValue As CTab)
    Set m_tab = newValue
End Property

Public Sub redrawTab()
    m_session.client.redrawSwitchbarTab m_tab
End Sub

Public Property Get session() As CSession
    Set session = m_session
End Property

Public Property Get channel() As CChannel
    Set channel = m_channel
End Property

Private Property Get ITabWindow_getTab() As CTab
    Set ITabWindow_getTab = m_tab
End Property

Private Sub ITextWindow_addEvent(eventName As String, params() As String)
    m_textView.addEvent eventName, params
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addEventEx(eventName As String, userStyle As CUserStyle, username As String, _
    flags As Long, params() As String)
    
    m_textView.addEventEx eventName, userStyle, username, flags, params
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addText(text As String)
    m_textView.addRawText "$0", makeStringArray(text)
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addTextEx(eventColour As CEventColour, foreColour As Byte, format As String, _
    userStyle As CUserStyle, username As String, flags As Long, params() As String)
    
    m_textView.addRawTextEx eventColour, foreColour, format, userStyle, username, flags, params
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
    fontsChanged
End Sub

Private Property Let ITextWindow_eventManager(RHS As CEventManager)
    m_textView.eventManager = RHS
End Property

Private Property Let ITextWindow_inputText(RHS As String)
    m_textInput.text = RHS
End Property

Private Property Get ITextWindow_session() As CSession
    Set ITextWindow_session = m_session
End Property

Private Sub ITextWindow_Update()
    m_textView.updateVisibility
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_nicklist_changeWidth(newWidth As Long)
    If newWidth < 5 Then
        m_nicklistWidth = 5
    ElseIf newWidth >= UserControl.ScaleWidth Then
        m_nicklistWidth = UserControl.ScaleWidth - 1
    Else
        m_nicklistWidth = newWidth
    End If
    
    UserControl_Resize
End Sub

Private Sub m_nicklist_doubleClicked(item As CNicklistItem)
    m_session.query item.text
End Sub

Private Sub m_nicklist_rightClicked(item As CNicklistItem)
    Set m_selectedNick = item
    
    If ignoreManager.isIgnored(m_session.getIal(m_selectedNick.text, ialAll), IGNORE_ANY) Then
        menuIgnore.Checked = True
    Else
        menuIgnore.Checked = False
    End If

    PopupMenu menuNicklist
End Sub

Private Sub m_textInput_tabbed(text As String, start As Long, Length As Long)
    Dim count As Long
    Dim matchLength As Long

    If m_lastTabIndex <> 0 Then
        If text = m_lastTab Then
            matchLength = Len(m_lastTabMatch)
            
            For count = m_lastTabIndex + 1 To m_nicklist.itemCount
                 If LCase$(left$(m_nicklist.item(count).text, matchLength)) = m_lastTabMatch Then
                    m_lastTabIndex = count
                    m_lastTab = m_nicklist.item(count).text
                    
                    m_textInput.replaceText start, Length, m_lastTab
                    Exit Sub
                 End If
            Next count
        
            If m_lastTabIndex > 1 Then
                For count = 1 To m_lastTabIndex - 1
                    If LCase$(left$(m_nicklist.item(count).text, matchLength)) = m_lastTabMatch Then
                        m_lastTabIndex = count
                        m_lastTab = m_nicklist.item(count).text
                        
                        m_textInput.replaceText start, Length, m_lastTab
                        Exit Sub
                    End If
                Next count
            End If
        
            Exit Sub
        End If
    End If

    If left$(text, 1) = "#" Then
        If Len(text) > 1 Then
            m_session.channelTabbing m_textInput, text, start, Length
        Else
            m_textInput.replaceText start, Length, m_channel.name
        End If
        
        Exit Sub
    End If
    
    m_lastTabMatch = LCase$(text)
    matchLength = Len(text)
    
    For count = 1 To m_nicklist.itemCount
        If LCase$(left$(m_nicklist.item(count).text, matchLength)) = m_lastTabMatch Then
            m_lastTabIndex = count
            m_lastTab = m_nicklist.item(count).text
            
            m_textInput.replaceText start, Length, m_lastTab
            
            Exit For
        End If
    Next count
End Sub

Private Sub m_textInput_textSubmitted(text As String, ctrl As Boolean)
    If left(LTrim$(text), 1) = "/" And Not ctrl Then
        m_session.textInput Me, text
    Else
        m_channel.textEntered text
    End If
End Sub

Private Sub m_textView_doubleClick()
    Dim channelCentral As frmChannelCentral
    
    Set channelCentral = New frmChannelCentral
    
    Dim fontUser As IFontUser
    
    Set fontUser = channelCentral
    fontUser.fontManager = m_fontManager
    
    channelCentral.channel = m_channel
    
    channelCentral.Show vbModal, Me
    Unload channelCentral
End Sub

Private Sub menuBan_Click()
    massBan
End Sub

Private Sub menuDehalfop_Click()
    massModeChange False, "h"
End Sub

Private Sub menuHalfop_Click()
    massModeChange True, "h"
End Sub

Private Sub menuHuggle_Click()
    m_channel.sendEmote "huggles " & m_selectedNick.text
End Sub

Private Sub menuIgnore_Click()
    Dim ignore As CIgnoreItem
    Dim mask As String
    
    mask = m_session.getIal(m_selectedNick.text, ialAll)

    If ignoreManager.isIgnored(mask, IGNORE_ANY) Then
        ignoreManager.removeIgnoreByMask mask
        ITextWindow_addEvent "IGNORE_REMOVED", makeStringArray(mask)
    Else
        Set ignore = New CIgnoreItem
        
        ignore.mask = m_session.getIal(m_selectedNick.text, ialHost)
        ignore.flags = IGNORE_ALL
        ignoreManager.addIgnore ignore
            
        ITextWindow_addEvent "IGNORE_ADDED", makeStringArray(ignore.mask, ignore.flagChars)
    End If
    
    saveIgnoreFile
End Sub

Private Sub menuKick_Click()
    massKick "No reason given"
End Sub

Private Sub menuKickBan_Click()
    massBan
    massKick "No reason given"
End Sub

Private Sub menuKickBanReason_Click()
    Dim result As Variant
    
    result = requestInput("Kickban with reason", "Enter a kickban reason", "Kickbanned", Me)
    
    If result = False Then
        Exit Sub
    End If
    
    massBan
    massKick result
End Sub

Private Sub menuKickReason_Click()
    Dim result As Variant
    
    result = requestInput("Kick with reason", "Enter a kick reason", "Kicked", Me)
    
    If result = False Then
        Exit Sub
    End If
    
    massKick result
End Sub

Private Sub menuOp_Click()
    massModeChange True, "o"
End Sub

Private Sub menuDeop_Click()
    massModeChange False, "o"
End Sub

Private Sub menuQuery_Click()
    m_session.query m_selectedNick.text
End Sub

Private Sub menuSlap_Click()
    m_channel.sendEmote "slaps " & m_selectedNick.text & " around a bit with a large trout"
End Sub

Private Sub menuVoice_Click()
    massModeChange True, "v"
End Sub

Private Sub menuDeVoice_Click()
    massModeChange False, "v"
End Sub

Private Sub massModeChange(op As Boolean, mode As String)
    Dim selectedItems As New cArrayList
    
    m_nicklist.getSelectedItems selectedItems
    
    If selectedItems.count = 0 Then
        Exit Sub
    End If
    
    Dim count As Long
    
    Dim modes As String
    Dim params As String
    
    If op Then
        modes = "+"
    Else
        modes = "-"
    End If
    
    For count = 1 To selectedItems.count
        modes = modes & mode
        params = params & selectedItems.item(count).text & " "
    Next count
    
    m_session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub massKick(ByVal reason As String)
    Dim selectedItems As New cArrayList
    
    m_nicklist.getSelectedItems selectedItems
    
    If selectedItems.count = 0 Then
        Exit Sub
    End If
    
    Dim count As Long

    For count = 1 To selectedItems.count
        m_session.sendLine "KICK " & m_channel.name & " " & selectedItems.item(count).text & " :" & reason
    Next count
End Sub

Private Sub massBan()
    Dim selectedItems As New cArrayList
    
    m_nicklist.getSelectedItems selectedItems
    
    If selectedItems.count = 0 Then
        Exit Sub
    End If
    
    Dim count As Long
    Dim modes As String
    Dim params As String
    
    modes = "+"
    
    For count = 1 To selectedItems.count
        modes = modes & "b"
        params = params & m_session.getIal(selectedItems.item(count).text, ialHost) & " "
    Next count

    m_session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub menuWhois_Click()
    m_session.sendLine "WHOIS " & m_selectedNick.text
End Sub

Private Sub mnuCtcpPing_Click()
    m_session.sendCtcp m_selectedNick.text, "PING"
End Sub

Private Sub mnuCtcpVersion_Click()
    m_session.sendCtcp m_selectedNick.text, "VERSION"
End Sub

Private Sub mnuCtcpTime_Click()
    m_session.sendCtcp m_selectedNick.text, "TIME"
End Sub

Private Sub UserControl_Initialize()
    m_nicklistWidth = 175
    m_textInputHeight = 30

    Dim window As IWindow

    Set m_nicklist = createNewWindow("swiftIrc.ctlNicklist", "nicklist")
    Set window = m_nicklist
    
    window.realWindow.visible = True
    
    Set m_textView = createNewWindow("swiftirc.ctlTextView", "textview")
    Set window = m_textView
    
    m_textView.ignoreSeperators = True
    
    window.realWindow.visible = True
    
    Set m_textInput = createNewWindow("swiftirc.ctlTextInput", "textInput")
    Set window = m_textInput
    
    window.realWindow.visible = True
End Sub

Private Sub fontsChanged()
    Dim count As Integer
    Dim fontUser As IFontUser
    
    m_textInputHeight = m_fontManager.fontHeight + 5
    
    For count = 1 To m_children.count
        If TypeOf m_children.item(count).object Is IFontUser Then
            Set fontUser = m_children.item(count).object
            fontUser.fontManager = m_fontManager
            fontUser.fontsUpdated
        End If
    Next count
    
    UserControl_Resize
End Sub

Private Function createNewWindow(progId As String, name As String) As IWindow
    Dim newControl As VBControlExtender
    
    Set newControl = Controls.Add(progId, name & m_nameCounter)
    
    Dim window As IWindow
    
    Set window = newControl.object
    window.realWindow = newControl
    
    m_children.Add newControl
    m_nameCounter = m_nameCounter + 1
    
    If TypeOf window Is IFontUser Then
        If Not m_fontManager Is Nothing Then
            Dim fontUser As IFontUser
            
            Set fontUser = window
            fontUser.fontManager = m_fontManager
            fontUser.fontsUpdated
        End If
    End If
    
    Set createNewWindow = window
End Function

Private Sub UserControl_Resize()
    Dim window As IWindow
    
    If UserControl.ScaleHeight <= m_textInputHeight Then
        Exit Sub
    End If
    
    If UserControl.ScaleWidth <= m_nicklistWidth Then
        Exit Sub
    End If
    
    If Not m_nicklist Is Nothing Then
        Set window = m_nicklist
        window.realWindow.Move UserControl.ScaleWidth - m_nicklistWidth, 0, m_nicklistWidth, _
            UserControl.ScaleHeight - m_textInputHeight
    End If
    
    If Not m_textView Is Nothing Then
        Set window = m_textView
        window.realWindow.Move 0, 0, UserControl.ScaleWidth - m_nicklistWidth, _
            UserControl.ScaleHeight - m_textInputHeight
    End If
    
    If Not m_textInput Is Nothing Then
        Set window = m_textInput
        window.realWindow.Move 0, UserControl.ScaleHeight - m_textInputHeight, _
            UserControl.ScaleWidth, m_textInputHeight
    End If
End Sub

Public Sub addNicklistItem(nick As String, prefix As String)
    m_nicklist.addItem nick, prefix
End Sub

Public Sub removeNicklistItem(nick As String, prefix As String)
    m_nicklist.removeItem nick, prefix
End Sub

Public Sub clearNicklist()
    m_nicklist.clearItems
End Sub

Private Sub UserControl_Terminate()
    debugLog "ctlWindowChannel terminated"
End Sub


