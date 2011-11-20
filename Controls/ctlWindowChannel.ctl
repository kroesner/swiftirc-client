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
         Begin VB.Menu menuIgnorePublic 
            Caption         =   "Public"
         End
         Begin VB.Menu menuIgnorePrivate 
            Caption         =   "Private"
         End
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
10        m_textView.pageDown
End Sub

Private Sub m_textInput_pageUp()
10        m_textView.pageUp
End Sub


Private Property Get ITextWindow_textview() As ctlTextView
10        Set ITextWindow_textview = m_textView
End Property

Private Sub ITextWindow_clear()
10        m_textView.clear
End Sub

Private Sub m_textView_clickedUrl(url As String)
10        If left$(url, 1) = "#" Then
20            If InStr(url, ",0") = 0 Then
30                m_session.joinChannel url
40            End If
50        Else
60            m_session.client.visitUrl url, True
70        End If
End Sub

Private Sub m_textView_noLongerNeedFocus()
10        If getRealWindow(m_textInput).visible Then
20            getRealWindow(m_textInput).setFocus
30        End If
End Sub

Private Sub m_textInput_mouseWheel(delta As Long)
10        m_textView.processMouseWheel delta
End Sub

Private Sub ITextWindow_focusInput()
10        If getRealWindow(m_textInput).visible Then
20            getRealWindow(m_textInput).setFocus
30        End If
End Sub

Private Sub IColourUser_coloursUpdated()
10        updateColours Controls
End Sub

Public Sub init(session As CSession, channel As CChannel)
10        Set m_session = session
20        Set m_channel = channel
30        m_nicklist.session = session
End Sub

Public Sub deInit()
10        Set m_session = Nothing
20        Set m_channel = Nothing
30        m_nicklist.session = Nothing
40        Set m_realWindow = Nothing
End Sub

Public Property Get switchbartab() As CTab
10        Set switchbartab = m_tab
End Property

Public Property Let switchbartab(newValue As CTab)
10        Set m_tab = newValue
End Property

Public Sub redrawTab()
10        m_session.client.redrawSwitchbarTab m_tab
End Sub

Public Property Get session() As CSession
10        Set session = m_session
End Property

Public Property Get channel() As CChannel
10        Set channel = m_channel
End Property

Private Property Get ITabWindow_getTab() As CTab
10        Set ITabWindow_getTab = m_tab
End Property

Private Sub ITextWindow_addEvent(eventName As String, params() As String)
10        m_textView.addEvent eventName, params
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addEventEx(eventName As String, userStyle As CUserStyle, username As String, _
    flags As Long, params() As String)
          
10        m_textView.addEventEx eventName, userStyle, username, flags, params
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addText(text As String)
10        m_textView.addRawText "$0", makeStringArray(text)
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addTextEx(eventColour As CEventColour, foreColour As Byte, format As String, _
    userStyle As CUserStyle, username As String, flags As Long, params() As String)
          
10        m_textView.addRawTextEx eventColour, foreColour, format, userStyle, username, flags, params
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
10        fontsChanged
End Sub

Private Property Let ITextWindow_eventManager(RHS As CEventManager)
10        m_textView.eventManager = RHS
End Property

Private Property Let ITextWindow_inputText(RHS As String)
10        m_textInput.text = RHS
End Property

Private Property Get ITextWindow_session() As CSession
10        Set ITextWindow_session = m_session
End Property

Private Sub ITextWindow_Update()
10        m_textView.updateVisibility
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_nicklist_changeWidth(newWidth As Long)
10        If newWidth < 5 Then
20            m_nicklistWidth = 5
30        ElseIf newWidth >= UserControl.ScaleWidth Then
40            m_nicklistWidth = UserControl.ScaleWidth - 1
50        Else
60            m_nicklistWidth = newWidth
70        End If
          
80        UserControl_Resize
End Sub

Private Sub m_nicklist_doubleClicked(item As CNicklistItem)
10        m_session.query item.text
End Sub

Private Sub m_nicklist_rightClicked(item As CNicklistItem)
10        Set m_selectedNick = item
          
20        If ignoreManager.isIgnored(m_session.getIal(m_selectedNick.text, ialAll), IGNORE_CHANNEL) Then
30            menuIgnorePublic.Checked = True
40        Else
50            menuIgnorePublic.Checked = False
60        End If
          
70        If ignoreManager.isIgnored(m_session.getIal(m_selectedNick.text, ialAll), IGNORE_PRIVATE) Then
80            menuIgnorePrivate.Checked = True
90        Else
100           menuIgnorePrivate.Checked = False
110       End If
          
120       PopupMenu menuNicklist
End Sub

Private Sub m_textInput_tabbed(text As String, start As Long, Length As Long)
          Dim count As Long
          Dim matchLength As Long

10        If m_lastTabIndex <> 0 Then
20            If text = m_lastTab Then
30                matchLength = Len(m_lastTabMatch)
                  
40                For count = m_lastTabIndex + 1 To m_nicklist.itemCount
50                     If LCase$(left$(m_nicklist.item(count).text, matchLength)) = m_lastTabMatch Then
60                        m_lastTabIndex = count
70                        m_lastTab = m_nicklist.item(count).text
                          
80                        m_textInput.replaceText start, Length, m_lastTab
90                        Exit Sub
100                    End If
110               Next count
              
120               If m_lastTabIndex > 1 Then
130                   For count = 1 To m_lastTabIndex - 1
140                       If LCase$(left$(m_nicklist.item(count).text, matchLength)) = m_lastTabMatch Then
150                           m_lastTabIndex = count
160                           m_lastTab = m_nicklist.item(count).text
                              
170                           m_textInput.replaceText start, Length, m_lastTab
180                           Exit Sub
190                       End If
200                   Next count
210               End If
              
220               Exit Sub
230           End If
240       End If

250       If left$(text, 1) = "#" Then
260           If Len(text) > 1 Then
270               m_session.channelTabbing m_textInput, text, start, Length
280           Else
290               m_textInput.replaceText start, Length, m_channel.name
300           End If
              
310           Exit Sub
320       End If
          
330       m_lastTabMatch = LCase$(text)
340       matchLength = Len(text)
          
350       For count = 1 To m_nicklist.itemCount
360           If LCase$(left$(m_nicklist.item(count).text, matchLength)) = m_lastTabMatch Then
370               m_lastTabIndex = count
380               m_lastTab = m_nicklist.item(count).text
                  
390               m_textInput.replaceText start, Length, m_lastTab
                  
400               Exit For
410           End If
420       Next count
End Sub

Private Sub m_textInput_textSubmitted(text As String, ctrl As Boolean)
10        If left(LTrim$(text), 1) = "/" And Not ctrl Then
20            m_session.textInput Me, text
30        Else
40            m_channel.textEntered text
50        End If
End Sub

Private Sub m_textView_doubleClick()
          Dim channelCentral As frmChannelCentral
          
10        Set channelCentral = New frmChannelCentral
          
          Dim fontUser As IFontUser
          
20        Set fontUser = channelCentral
30        fontUser.fontManager = m_fontManager
          
40        channelCentral.channel = m_channel
          
50        channelCentral.Show vbModal, Me
60        Unload channelCentral
End Sub

Private Sub menuBan_Click()
10        massBan
End Sub

Private Sub menuDehalfop_Click()
10        massModeChange False, "h"
End Sub

Private Sub menuHalfop_Click()
10        massModeChange True, "h"
End Sub

Private Sub menuHuggle_Click()
10        m_channel.sendEmote "huggles " & m_selectedNick.text
End Sub

Private Sub menuIgnorePrivate_Click()
          Dim ignore As CIgnoreItem
          Dim mask As String
          
10        mask = m_session.getIal(m_selectedNick.text, ialAll)

20        If ignoreManager.isIgnored(mask, IGNORE_PRIVATE) Then
30            Set ignore = ignoreManager.getIgnoreByMask(mask)
              
40            If Not ignore Is Nothing Then
50                ignore.flags = ignore.flags And (ALL_BITS - IGNORE_PRIVATE_EXTENDED)
                  
60                If ignore.flags = IGNORE_NONE Then
70                    mask = ignore.mask
80                    ignoreManager.removeIgnoreByMask mask
90                    ITextWindow_addEvent "IGNORE_REMOVED", makeStringArray(mask)
100               Else
110                   ITextWindow_addEvent "IGNORE_UPDATED", makeStringArray(ignore.mask, ignore.flagChars)
120               End If
130           End If
140       Else
150           Set ignore = ignoreManager.getIgnoreByMask(m_session.getIal(m_selectedNick.text, ialAll))
              
160           If Not ignore Is Nothing Then
170               ignore.flags = ignore.flags Or IGNORE_PRIVATE_EXTENDED
180               ITextWindow_addEvent "IGNORE_UPDATED", makeStringArray(ignore.mask, ignore.flagChars)
190           Else
200               Set ignore = New CIgnoreItem
              
210               ignore.mask = m_session.getIal(m_selectedNick.text, ialHost)
220               ignore.flags = IGNORE_PRIVATE_EXTENDED
230               ignoreManager.addIgnore ignore
                  
240               ITextWindow_addEvent "IGNORE_ADDED", makeStringArray(ignore.mask, ignore.flagChars)
250           End If
260       End If
          
270       saveIgnoreFile
End Sub

Private Sub menuIgnorePublic_Click()
          Dim ignore As CIgnoreItem
          Dim mask As String
          
10        mask = m_session.getIal(m_selectedNick.text, ialAll)

20        If ignoreManager.isIgnored(mask, IGNORE_CHANNEL) Then
30            Set ignore = ignoreManager.getIgnoreByMask(mask)
              
40            If Not ignore Is Nothing Then
50                ignore.flags = ignore.flags And (ALL_BITS - IGNORE_CHANNEL)
                  
60                If ignore.flags = IGNORE_NONE Then
70                    mask = ignore.mask
80                    ignoreManager.removeIgnoreByMask mask
90                    ITextWindow_addEvent "IGNORE_REMOVED", makeStringArray(mask)
100               Else
110                   ITextWindow_addEvent "IGNORE_UPDATED", makeStringArray(ignore.mask, ignore.flagChars)
120               End If
130           End If
140       Else
150           Set ignore = ignoreManager.getIgnoreByMask(mask)
              
160           If Not ignore Is Nothing Then
170               ignore.flags = ignore.flags Or IGNORE_CHANNEL
180               ITextWindow_addEvent "IGNORE_UPDATED", makeStringArray(ignore.mask, ignore.flagChars)
190           Else
200               Set ignore = New CIgnoreItem
              
210               ignore.mask = m_session.getIal(m_selectedNick.text, ialHost)
220               ignore.flags = IGNORE_CHANNEL
230               ignoreManager.addIgnore ignore
                  
240               ITextWindow_addEvent "IGNORE_ADDED", makeStringArray(ignore.mask, ignore.flagChars)
250           End If
260       End If
          
270       saveIgnoreFile
End Sub

Private Sub menuKick_Click()
10        massKick "No reason given"
End Sub

Private Sub menuKickBan_Click()
10        massBan
20        massKick "No reason given"
End Sub

Private Sub menuKickBanReason_Click()
          Dim result As Variant
          
10        result = requestInput("Kickban with reason", "Enter a kickban reason", "Kickbanned", Me)
          
20        If result = False Then
30            Exit Sub
40        End If
          
50        massBan
60        massKick result
End Sub

Private Sub menuKickReason_Click()
          Dim result As Variant
          
10        result = requestInput("Kick with reason", "Enter a kick reason", "Kicked", Me)
          
20        If result = False Then
30            Exit Sub
40        End If
          
50        massKick result
End Sub

Private Sub menuOp_Click()
10        massModeChange True, "o"
End Sub

Private Sub menuDeop_Click()
10        massModeChange False, "o"
End Sub

Private Sub menuQuery_Click()
10        m_session.query m_selectedNick.text
End Sub

Private Sub menuSlap_Click()
10        m_channel.sendEmote "slaps " & m_selectedNick.text & " around a bit with a large trout"
End Sub

Private Sub menuVoice_Click()
10        massModeChange True, "v"
End Sub

Private Sub menuDeVoice_Click()
10        massModeChange False, "v"
End Sub

Private Sub massModeChange(op As Boolean, mode As String)
          Dim selectedItems As New cArrayList
          
10        m_nicklist.getSelectedItems selectedItems
          
20        If selectedItems.count = 0 Then
30            Exit Sub
40        End If
          
          Dim count As Long
          
          Dim modes As String
          Dim params As String
          
50        If op Then
60            modes = "+"
70        Else
80            modes = "-"
90        End If
          
100       For count = 1 To selectedItems.count
110           modes = modes & mode
120           params = params & selectedItems.item(count).text & " "
130       Next count
          
140       m_session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub massKick(ByVal reason As String)
          Dim selectedItems As New cArrayList
          
10        m_nicklist.getSelectedItems selectedItems
          
20        If selectedItems.count = 0 Then
30            Exit Sub
40        End If
          
          Dim count As Long

50        For count = 1 To selectedItems.count
60            m_session.sendLine "KICK " & m_channel.name & " " & selectedItems.item(count).text & " :" & reason
70        Next count
End Sub

Private Sub massBan()
          Dim selectedItems As New cArrayList
          
10        m_nicklist.getSelectedItems selectedItems
          
20        If selectedItems.count = 0 Then
30            Exit Sub
40        End If
          
          Dim count As Long
          Dim modes As String
          Dim params As String
          
50        modes = "+"
          
60        For count = 1 To selectedItems.count
70            modes = modes & "b"
80            params = params & m_session.getIal(selectedItems.item(count).text, ialHost) & " "
90        Next count

100       m_session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub menuWhois_Click()
10        m_session.sendLine "WHOIS " & m_selectedNick.text
End Sub

Private Sub mnuCtcpPing_Click()
10        m_session.sendCtcp m_selectedNick.text, "PING"
End Sub

Private Sub mnuCtcpVersion_Click()
10        m_session.sendCtcp m_selectedNick.text, "VERSION"
End Sub

Private Sub mnuCtcpTime_Click()
10        m_session.sendCtcp m_selectedNick.text, "TIME"
End Sub

Private Sub UserControl_Initialize()
10        m_nicklistWidth = 175
20        m_textInputHeight = 30

          Dim window As IWindow

30        Set m_nicklist = createNewWindow("swiftIrc.ctlNicklist", "nicklist")
40        Set window = m_nicklist
          
50        window.realWindow.visible = True
          
60        Set m_textView = createNewWindow("swiftirc.ctlTextView", "textview")
70        Set window = m_textView
          
80        m_textView.ignoreSeperators = True
          
90        window.realWindow.visible = True
          
100       Set m_textInput = createNewWindow("swiftirc.ctlTextInput", "textInput")
110       Set window = m_textInput
          
120       window.realWindow.visible = True
End Sub

Private Sub fontsChanged()
          Dim count As Integer
          Dim fontUser As IFontUser
          
10        m_textInputHeight = m_fontManager.fontHeight + 5
          
20        For count = 1 To m_children.count
30            If TypeOf m_children.item(count).object Is IFontUser Then
40                Set fontUser = m_children.item(count).object
50                fontUser.fontManager = m_fontManager
60                fontUser.fontsUpdated
70            End If
80        Next count
          
90        UserControl_Resize
End Sub

Private Function createNewWindow(progId As String, name As String) As IWindow
          Dim newControl As VBControlExtender
          
10        Set newControl = Controls.Add(progId, name & m_nameCounter)
          
          Dim window As IWindow
          
20        Set window = newControl.object
30        window.realWindow = newControl
          
40        m_children.Add newControl
50        m_nameCounter = m_nameCounter + 1
          
60        If TypeOf window Is IFontUser Then
70            If Not m_fontManager Is Nothing Then
                  Dim fontUser As IFontUser
                  
80                Set fontUser = window
90                fontUser.fontManager = m_fontManager
100               fontUser.fontsUpdated
110           End If
120       End If
          
130       Set createNewWindow = window
End Function

Private Sub UserControl_Resize()
          Dim window As IWindow
          
10        If UserControl.ScaleHeight <= m_textInputHeight Then
20            Exit Sub
30        End If
          
40        If UserControl.ScaleWidth <= m_nicklistWidth Then
50            Exit Sub
60        End If
          
70        If Not m_nicklist Is Nothing Then
80            Set window = m_nicklist
90            window.realWindow.Move UserControl.ScaleWidth - m_nicklistWidth, 0, m_nicklistWidth, _
                  UserControl.ScaleHeight - m_textInputHeight
100       End If
          
110       If Not m_textView Is Nothing Then
120           Set window = m_textView
130           window.realWindow.Move 0, 0, UserControl.ScaleWidth - m_nicklistWidth, _
                  UserControl.ScaleHeight - m_textInputHeight
140       End If
          
150       If Not m_textInput Is Nothing Then
160           Set window = m_textInput
170           window.realWindow.Move 0, UserControl.ScaleHeight - m_textInputHeight, _
                  UserControl.ScaleWidth, m_textInputHeight
180       End If
End Sub

Public Sub addNicklistItem(nick As String, prefix As String)
10        m_nicklist.addItem nick, prefix
End Sub

Public Sub removeNicklistItem(nick As String, prefix As String)
10        m_nicklist.removeItem nick, prefix
End Sub

Public Sub clearNicklist()
10        m_nicklist.clearItems
End Sub

Private Sub UserControl_Terminate()
10        debugLog "ctlWindowChannel terminated"
End Sub


