VERSION 5.00
Begin VB.UserControl ctlOptionsColour 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.ComboBox comboSwitchbarRows 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1230
      Width           =   2700
   End
   Begin VB.ComboBox comboSwitchbarPosition 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   765
      Width           =   2685
   End
   Begin VB.ComboBox comboTheme 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   2715
   End
End
Attribute VB_Name = "ctlOptionsColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IColourUser
Implements ISubclass

Private m_labelManager As New CLabelManager
Private m_realWindow As VBControlExtender

Private WithEvents m_ctlEventColourEditor As swiftIrc.ctlEventColourEditor
Attribute m_ctlEventColourEditor.VB_VarHelpID = -1
Private WithEvents m_ctlColourPalette As swiftIrc.ctlColourPalette
Attribute m_ctlColourPalette.VB_VarHelpID = -1
Private WithEvents m_buttonNewTheme As swiftIrc.ctlButton
Attribute m_buttonNewTheme.VB_VarHelpID = -1
Private WithEvents m_buttonDeleteTheme As swiftIrc.ctlButton
Attribute m_buttonDeleteTheme.VB_VarHelpID = -1

Private WithEvents m_buttonChangeFont As swiftIrc.ctlButton
Attribute m_buttonChangeFont.VB_VarHelpID = -1
Private WithEvents m_buttonDefaultFont As swiftIrc.ctlButton
Attribute m_buttonDefaultFont.VB_VarHelpID = -1

Private m_labelFontInfo As CLabel

Private m_client As swiftIrc.SwiftIrcClient

Private m_fontName As String
Private m_fontSize As Integer
Private m_fontBold As Boolean
Private m_fontItalic As Boolean

Private m_fontManager As CFontManager

Private m_buttonSbEventColour As swiftIrc.ctlButton
Private m_buttonSbMessageColour As swiftIrc.ctlButton
Private m_buttonSbAlertColour As swiftIrc.ctlButton
Private m_buttonSbHighlightColour As swiftIrc.ctlButton

Private m_checkSbEventFlash As VB.CheckBox
Private m_checkSbMessageFlash As VB.CheckBox
Private m_checkSbAlertFlash As VB.CheckBox
Private m_checkSbHighlightFlash As VB.CheckBox

Private WithEvents m_colourEvent As swiftIrc.ctlSingleColourSelector
Attribute m_colourEvent.VB_VarHelpID = -1
Private WithEvents m_colourMessage As swiftIrc.ctlSingleColourSelector
Attribute m_colourMessage.VB_VarHelpID = -1
Private WithEvents m_colourAlert As swiftIrc.ctlSingleColourSelector
Attribute m_colourAlert.VB_VarHelpID = -1
Private WithEvents m_colourHighlight As swiftIrc.ctlSingleColourSelector
Attribute m_colourHighlight.VB_VarHelpID = -1

Private m_sbTabEvent As CTab
Private m_sbTabMessage As CTab
Private m_sbTabAlert As CTab
Private m_sbTabHighlight As CTab

Private m_transTimer As Long
Private m_trans As Boolean

Private m_themes As New Collection
Private m_currentTheme As CColourTheme

Public Property Get client() As swiftIrc.SwiftIrcClient
10        Set client = m_client
End Property

Public Property Let client(newValue As swiftIrc.SwiftIrcClient)
10        Set m_client = newValue
End Property

Private Sub comboTheme_Click()
10        Set m_currentTheme = m_themes.item(comboTheme.ListIndex + 1)
20        loadTheme m_currentTheme
End Sub

Private Sub IColourUser_coloursUpdated()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
20        updateColours Controls
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        Select Case CurrentMessage
              Case WM_TIMER
20                ISubclass_MsgResponse = emrConsume
30        End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
10        Select Case iMsg
              Case WM_TIMER
20                If wParam = m_transTimer Then
30                    m_trans = Not m_trans
                      
40                    If m_checkSbEventFlash.value Then
50                        m_sbTabEvent.trans = m_trans
60                    Else
70                        m_sbTabEvent.trans = False
80                    End If
                      
90                    If m_checkSbMessageFlash.value Then
100                       m_sbTabMessage.trans = m_trans
110                   Else
120                       m_sbTabMessage.trans = False
130                   End If
                      
140                   If m_checkSbAlertFlash.value Then
150                       m_sbTabAlert.trans = m_trans
160                   Else
170                       m_sbTabAlert.trans = False
180                   End If
                      
190                   If m_checkSbHighlightFlash.value Then
200                       m_sbTabHighlight.trans = m_trans
210                   Else
220                       m_sbTabHighlight.trans = False
230                   End If
                      
240                   UserControl_Paint
250               End If
260       End Select
End Function

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Sub initControls()
10        m_labelManager.addLabel "Colour theme", ltHeading, 10, 10
20        m_labelManager.addLabel "Font", ltHeading, 320, 10
30        m_labelManager.addLabel "Switchbar", ltHeading, 320, 90
          
40        Set m_sbTabEvent = New CTab
50        Set m_sbTabMessage = New CTab
60        Set m_sbTabAlert = New CTab
70        Set m_sbTabHighlight = New CTab
          
80        m_sbTabEvent.icon = g_iconSBChannel
90        m_sbTabEvent.foreColour = getPaletteEntry(settings.setting("switchbarColourEvent", estNumber))
100       m_sbTabEvent.caption = "Events"

110       m_sbTabMessage.icon = g_iconSBChannel
120       m_sbTabMessage.foreColour = getPaletteEntry(settings.setting("switchbarColourMessage", estNumber))
130       m_sbTabMessage.caption = "Messages"
          
140       m_sbTabAlert.icon = g_iconSBQuery
150       m_sbTabAlert.foreColour = getPaletteEntry(settings.setting("switchbarColourAlert", estNumber))
160       m_sbTabAlert.caption = "Alerts (PMs)"
          
170       m_sbTabHighlight.icon = g_iconSBChannel
180       m_sbTabHighlight.foreColour = getPaletteEntry(settings.setting("switchbarColourHighlight", estNumber))
190       m_sbTabHighlight.caption = "Highlights"
          
200       Set m_colourEvent = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
210       Set m_colourMessage = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
220       Set m_colourAlert = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
230       Set m_colourHighlight = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
          
240       m_colourEvent.colour = settings.setting("switchbarColourEvent", eSettingType.estNumber)
250       m_colourMessage.colour = settings.setting("switchbarColourMessage", eSettingType.estNumber)
260       m_colourAlert.colour = settings.setting("switchbarColourAlert", eSettingType.estNumber)
270       m_colourHighlight.colour = settings.setting("switchbarColourHighlight", eSettingType.estNumber)
          
280       getRealWindow(m_colourEvent).Move 425, 115, 20, 20
290       getRealWindow(m_colourMessage).Move 425, 140, 20, 20
300       getRealWindow(m_colourAlert).Move 425, 165, 20, 20
310       getRealWindow(m_colourHighlight).Move 425, 190, 20, 20
          
320       Set m_checkSbEventFlash = addCheckBox(Controls, "Flash icon", 450, 115, 100, 20)
330       Set m_checkSbMessageFlash = addCheckBox(Controls, "Flash icon", 450, 140, 100, 20)
340       Set m_checkSbAlertFlash = addCheckBox(Controls, "Flash icon", 450, 165, 100, 20)
350       Set m_checkSbHighlightFlash = addCheckBox(Controls, "Flash icon", 450, 190, 100, 20)
          
360       m_checkSbEventFlash.value = -settings.setting("switchbarFlashEvent", eSettingType.estBoolean)
370       m_checkSbMessageFlash.value = -settings.setting("switchbarFlashMessage", eSettingType.estBoolean)
380       m_checkSbAlertFlash.value = -settings.setting("switchbarFlashAlert", eSettingType.estBoolean)
390       m_checkSbHighlightFlash.value = -settings.setting("switchbarFlashHighlight", eSettingType.estBoolean)
          
400       m_labelManager.addLabel "Switchbar position:", ltNormal, 320, 220
410       m_labelManager.addLabel "Switchbar rows:", ltNormal, 320, 270
          
420       comboSwitchbarPosition.addItem "Top"
430       comboSwitchbarPosition.addItem "Bottom"
          
440       If StrComp(settings.setting("switchbarPosition", estString), "bottom", vbTextCompare) = 0 Then
450           comboSwitchbarPosition.ListIndex = 1
460       Else
470           comboSwitchbarPosition.ListIndex = 0
480       End If
          
          Dim count As Long
          
490       For count = 1 To 10
500           comboSwitchbarRows.addItem count
510       Next count
          
          Dim index As Long
          
520       index = settings.setting("switchbarRows", estNumber) - 1
          
530       If index >= 0 And index < comboSwitchbarRows.ListCount Then
540           comboSwitchbarRows.ListIndex = index
550       Else
560           comboSwitchbarRows.ListIndex = 0
570       End If
          
580       Set m_labelFontInfo = m_labelManager.addLabel("Current font: " & m_fontName & " size " & _
              CStr(m_fontSize), ltNormal, 320, 35)

590       Set m_ctlEventColourEditor = createControl(Controls, "swiftIrc.ctlEventColourEditor", _
              "eventColourEditor")
600       Set m_ctlColourPalette = createControl(Controls, "swiftIrc.ctlColourPalette", "colourPalette")
          
610       getRealWindow(m_ctlEventColourEditor).Move 10, 100, 280, 155
620       getRealWindow(m_ctlColourPalette).Move 10, 260, 200, 50
          
630       m_ctlColourPalette.allowPaletteChange = True
          
640       Set m_buttonNewTheme = addButton(Controls, "N&ew", 10, 35, 75, 20)
650       Set m_buttonDeleteTheme = addButton(Controls, "&Delete", 90, 35, 75, 20)
          
660       Set m_buttonChangeFont = addButton(Controls, "C&hange font", 360, 55, 100, 20)
670       Set m_buttonDefaultFont = addButton(Controls, "&Restore default", 465, 55, 100, 20)
End Sub

Private Sub m_buttonChangeFont_clicked()
          Dim cf As tChooseFont
          
10        cf.lStructSize = Len(cf)
20        cf.hwndOwner = UserControl.hwnd
30        cf.flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
40        cf.nSizeMin = 8
50        cf.nSizeMax = 28
          
          Dim lf As LOGFONT
          Dim result As Long
          
60        lf = m_fontManager.fontStruct
          
70        cf.lpLogFont = VarPtr(lf)
          
80        result = ChooseFont(cf)
          
90        If result = 0 Then
100           Exit Sub
110       End If
          
120       m_fontName = StrConv(lf.lfFaceName, vbUnicode)
130       m_fontSize = cf.iPointSize / 10
          
          Dim count As Long
          
140       For count = 1 To Len(m_fontName)
150           If Mid$(m_fontName, count, 1) = Chr$(0) Then
160               m_fontName = Mid$(m_fontName, 1, count - 1)
170           End If
180       Next count
          
190       If lf.lfWeight = FW_BOLD Then
200           m_fontBold = True
210       Else
220           m_fontBold = False
230       End If
          
240       If lf.lfItalic <> 0 Then
250           m_fontItalic = True
260       Else
270           m_fontItalic = False
280       End If
          
290       m_fontManager.changeFont UserControl.hdc, m_fontName, m_fontSize, m_fontBold, m_fontItalic
300       m_labelFontInfo.caption = "Current font: " & m_fontName & " size " & _
              CStr(m_fontSize)
310       UserControl_Paint
End Sub

Private Sub m_buttonDefaultFont_clicked()
10        m_fontName = getBestDefaultFont
20        m_fontSize = 9
30        m_fontBold = False
40        m_fontItalic = False
          
50        m_labelFontInfo.caption = "Current font: " & m_fontName & " size " & _
              m_fontSize
              
60        UserControl_Paint
End Sub

Private Sub m_buttonDeleteTheme_clicked()
10        If comboTheme.ListIndex = -1 Then
20            Exit Sub
30        End If

40        If m_themes.count = 1 Then
50            Exit Sub
60        End If
          
          Dim index As Long
          
70        index = comboTheme.ListIndex
          
80        m_themes.Remove comboTheme.ListIndex + 1
90        comboTheme.removeItem comboTheme.ListIndex
          
100       If index > 0 Then
110           comboTheme.ListIndex = index - 1
120       Else
130           comboTheme.ListIndex = 0
140       End If

150       Set m_currentTheme = m_themes.item(comboTheme.ListIndex + 1)
160       loadTheme m_currentTheme
End Sub

Private Sub m_buttonNewTheme_clicked()
          Dim theme As CColourTheme
          Dim result As Variant
          Dim name As String
          
10        result = requestInput("New theme", "Enter new theme name:", vbNullString, Me)
          
20        If result = False Then
30            Exit Sub
40        End If
          
50        name = result
          
60        If LenB(name) = 0 Then
70            Exit Sub
80        End If
          
90        If Not findTheme(name) Is Nothing Then
100           MsgBox "A theme already exists by this name", vbCritical, "Could not create theme"
110           Exit Sub
120       End If
          
130       Set theme = New CColourTheme
          
140       m_currentTheme.copy theme
          
150       theme.name = name
160       m_themes.Add theme, LCase$(name)
          
170       Set m_currentTheme = theme
          
180       comboTheme.addItem theme.name
190       comboTheme.ListIndex = comboTheme.ListCount - 1
End Sub

Private Function findTheme(name As String) As CColourTheme
10        On Error Resume Next
20        Set findTheme = m_themes.item(LCase$(name))
End Function

Private Sub m_colourMessage_colourChanged()
10        m_sbTabMessage.foreColour = m_currentTheme.paletteEntry(m_colourMessage.colour)
20        UserControl_Paint
End Sub

Private Sub m_colourHighlight_colourChanged()
10        m_sbTabHighlight.foreColour = m_currentTheme.paletteEntry(m_colourHighlight.colour)
20        UserControl_Paint
End Sub

Private Sub m_colourAlert_colourChanged()
10        m_sbTabAlert.foreColour = m_currentTheme.paletteEntry(m_colourAlert.colour)
20        UserControl_Paint
End Sub

Private Sub m_colourEvent_colourChanged()
10        m_sbTabEvent.foreColour = m_currentTheme.paletteEntry(m_colourEvent.colour)
20        UserControl_Paint
End Sub

Private Sub m_ctlColourPalette_colourChanged(index As Long, colour As Long)
10        m_ctlEventColourEditor.paletteEntryUpdated index, colour
20        m_currentTheme.paletteEntry(index) = colour
          
30        m_colourEvent.setPalette m_currentTheme.getPalette
40        m_colourMessage.setPalette m_currentTheme.getPalette
50        m_colourAlert.setPalette m_currentTheme.getPalette
60        m_colourHighlight.setPalette m_currentTheme.getPalette
End Sub

Private Sub m_ctlColourPalette_colourSelected(index As Long)
10        m_ctlEventColourEditor.colourSelected index
          
20        If m_ctlEventColourEditor.backgroundSelected Then
30            m_currentTheme.backgroundColour = index
40        Else
50            m_currentTheme.eventColour(m_ctlEventColourEditor.selectedItem) = index
60        End If
End Sub

Private Sub m_ctlEventColourEditor_itemClicked(colourIndex As Byte)
10        m_ctlColourPalette.setFocus colourIndex
End Sub

Private Sub UserControl_Initialize()
10        m_transTimer = 1
          
20        SetTimer UserControl.hwnd, m_transTimer, 500, 0
30        AttachMessage Me, UserControl.hwnd, WM_TIMER

40        Set m_fontManager = New CFontManager
          
50        m_fontName = settings.fontName
60        m_fontSize = settings.fontSize
70        m_fontBold = settings.setting("fontBold", estBoolean)
80        m_fontItalic = settings.setting("fontItalic", estBoolean)
          
90        m_fontManager.changeFont UserControl.hdc, m_fontName, m_fontSize, m_fontBold, m_fontItalic

100       UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
110       initControls
          
          Dim count As Long
          Dim index As Long
          Dim theme As CColourTheme
          
120       colourThemes.copyThemes m_themes
          
130       For Each theme In m_themes
140           comboTheme.addItem theme.name
              
150           count = count + 1
              
160           If index = 0 Then
170               If theme.name = colourThemes.currentTheme.name Then
180                   index = count
190               End If
200           End If
210       Next theme
          
220       Set m_currentTheme = m_themes.item(index)
230       comboTheme.ListIndex = index - 1
          
240       loadTheme m_currentTheme
End Sub

Private Sub loadTheme(theme As CColourTheme)
10        colourThemes.currentSettingsTheme = theme

20        m_ctlEventColourEditor.backgroundColour = theme.backgroundColour
30        m_ctlEventColourEditor.setPalette theme.getPalette()
40        m_ctlEventColourEditor.setEventColours theme.getEventColours()
50        m_ctlColourPalette.setPalette theme.getPalette()
          
60        m_colourEvent.setPalette theme.getPalette
70        m_colourMessage.setPalette theme.getPalette
80        m_colourAlert.setPalette theme.getPalette
90        m_colourHighlight.setPalette theme.getPalette
End Sub

Public Sub saveSettings()
10        colourThemes.clear
          
          Dim theme As CColourTheme
          
20        For Each theme In m_themes
30            colourThemes.addThemeIndirect theme
40        Next theme
          
50        colourThemes.currentTheme = m_currentTheme
60        colourThemes.saveThemes g_userPath & "\swiftirc_themes.xml"
          
70        eventColours.loadTheme colourThemes.currentTheme
80        g_textViewBack = colourThemes.currentTheme.backgroundColour
90        g_textViewFore = eventColours.normalText.colour
100       g_textInputBack = colourThemes.currentTheme.backgroundColour
110       g_textInputFore = eventColours.normalText.colour
120       g_nicklistBack = colourThemes.currentTheme.backgroundColour
130       g_nicklistFore = eventColours.normalText.colour
140       g_channelListBack = colourThemes.currentTheme.backgroundColour
150       g_channelListFore = eventColours.normalText.colour
          
160       settings.setting("switchbarColourEvent", estNumber) = m_colourEvent.colour
170       settings.setting("switchbarColourMessage", estNumber) = m_colourMessage.colour
180       settings.setting("switchbarColourAlert", estNumber) = m_colourAlert.colour
190       settings.setting("switchbarColourHighlight", estNumber) = m_colourHighlight.colour
          
200       settings.setting("switchbarFlashEvent", estBoolean) = -m_checkSbEventFlash.value
210       settings.setting("switchbarFlashMessage", estBoolean) = -m_checkSbMessageFlash.value
220       settings.setting("switchbarFlashAlert", estBoolean) = -m_checkSbAlertFlash.value
230       settings.setting("switchbarFlashHighlight", estBoolean) = -m_checkSbHighlightFlash.value
          
240       If comboSwitchbarPosition.ListIndex = 0 Then
250           settings.setting("switchbarPosition", estString) = "Top"
260       Else
270           settings.setting("switchbarPosition", estString) = "Bottom"
280       End If
          
290       settings.setting("switchbarRows", estNumber) = comboSwitchbarRows.ListIndex + 1
          
300       m_client.coloursUpdated
          
310       settings.fontName = m_fontName
320       settings.fontSize = m_fontSize
330       settings.setting("fontBold", estBoolean) = m_fontBold
340       settings.setting("fontItalic", estBoolean) = m_fontItalic
          
350       m_client.changeFont m_fontName, m_fontSize, m_fontBold, m_fontItalic
          
360       If StrComp(settings.setting("switchbarPosition", estString), "Top", vbTextCompare) = 0 Then
370           m_client.switchbar.position = sbpTop
380       ElseIf StrComp(settings.setting("switchbarPosition", estString), "Bottom", vbTextCompare) = 0 Then
390           m_client.switchbar.position = sbpBottom
400       End If
          
410       m_client.switchbar.rows = settings.setting("switchbarRows", estNumber)
End Sub

Private Sub UserControl_Paint()
          Dim backBuffer As Long
          Dim backBitmap As Long
          
          Dim oldBitmap As Long
          Dim oldFont As Long
          
10        backBuffer = CreateCompatibleDC(UserControl.hdc)
20        backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
          
30        oldBitmap = SelectObject(backBuffer, backBitmap)
40        SetBkColor backBuffer, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
          
50        oldFont = SelectObject(backBuffer, GetCurrentObject(UserControl.hdc, OBJ_FONT))

60        FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
          
70        m_labelManager.renderLabels backBuffer
          
80        FrameRect backBuffer, makeRect(5, 305, 5, UserControl.ScaleHeight - 5), _
              colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
90        FrameRect backBuffer, makeRect(315, UserControl.ScaleWidth - 5, 5, 80), _
              colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
          
100       FrameRect backBuffer, makeRect(315, UserControl.ScaleWidth - 5, 85, UserControl.ScaleHeight - 5), _
              colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
          
110       m_sbTabEvent.render backBuffer, 320, 115, 100, 20
120       m_sbTabMessage.render backBuffer, 320, 140, 100, 20
130       m_sbTabAlert.render backBuffer, 320, 165, 100, 20
140       m_sbTabHighlight.render backBuffer, 320, 190, 100, 20
          
150       BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
              backBuffer, 0, 0, vbSrcCopy
          
160       SelectObject backBuffer, oldBitmap
170       SelectObject backBuffer, oldFont
          
180       DeleteDC backBuffer
190       DeleteObject backBitmap
End Sub

Private Sub UserControl_Resize()
10        comboTheme.left = 10
20        comboTheme.top = 65
30        comboTheme.width = 280
          
40        comboSwitchbarPosition.left = 320
50        comboSwitchbarPosition.top = 240
60        comboSwitchbarPosition.width = 75
          
70        comboSwitchbarRows.left = 320
80        comboSwitchbarRows.top = 290
90        comboSwitchbarRows.width = 75
End Sub

Private Sub UserControl_Terminate()
10        DetachMessage Me, UserControl.hwnd, WM_TIMER
20        debugLog "ctlOptionsColour terminating"
End Sub
