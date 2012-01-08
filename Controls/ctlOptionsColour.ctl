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

Private m_fontName As String
Private m_fontSize As Integer
Private m_fontBold As Boolean
Private m_fontItalic As Boolean

Private m_fontmanager As CFontManager

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

Private Sub comboTheme_Click()
    Set m_currentTheme = m_themes.item(comboTheme.ListIndex + 1)
    loadTheme m_currentTheme
End Sub

Private Sub IColourUser_coloursUpdated()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    updateColours Controls
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_TIMER
            ISubclass_MsgResponse = emrConsume
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
        Case WM_TIMER
            If wParam = m_transTimer Then
                m_trans = Not m_trans
                
                If m_checkSbEventFlash.value Then
                    m_sbTabEvent.trans = m_trans
                Else
                    m_sbTabEvent.trans = False
                End If
                
                If m_checkSbMessageFlash.value Then
                    m_sbTabMessage.trans = m_trans
                Else
                    m_sbTabMessage.trans = False
                End If
                
                If m_checkSbAlertFlash.value Then
                    m_sbTabAlert.trans = m_trans
                Else
                    m_sbTabAlert.trans = False
                End If
                
                If m_checkSbHighlightFlash.value Then
                    m_sbTabHighlight.trans = m_trans
                Else
                    m_sbTabHighlight.trans = False
                End If
                
                UserControl_Paint
            End If
    End Select
End Function

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Sub initControls()
    m_labelManager.addLabel "Colour theme", ltHeading, 10, 10
    m_labelManager.addLabel "Font", ltHeading, 320, 10
    m_labelManager.addLabel "Switchbar", ltHeading, 320, 90
    
    Set m_sbTabEvent = New CTab
    Set m_sbTabMessage = New CTab
    Set m_sbTabAlert = New CTab
    Set m_sbTabHighlight = New CTab
    
    m_sbTabEvent.icon = g_iconSBChannel
    m_sbTabEvent.foreColour = getPaletteEntry(settings.setting("switchbarColourEvent", estNumber))
    m_sbTabEvent.caption = "Events"

    m_sbTabMessage.icon = g_iconSBChannel
    m_sbTabMessage.foreColour = getPaletteEntry(settings.setting("switchbarColourMessage", estNumber))
    m_sbTabMessage.caption = "Messages"
    
    m_sbTabAlert.icon = g_iconSBQuery
    m_sbTabAlert.foreColour = getPaletteEntry(settings.setting("switchbarColourAlert", estNumber))
    m_sbTabAlert.caption = "Alerts (PMs)"
    
    m_sbTabHighlight.icon = g_iconSBChannel
    m_sbTabHighlight.foreColour = getPaletteEntry(settings.setting("switchbarColourHighlight", estNumber))
    m_sbTabHighlight.caption = "Highlights"
    
    Set m_colourEvent = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
    Set m_colourMessage = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
    Set m_colourAlert = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
    Set m_colourHighlight = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
    
    m_colourEvent.colour = settings.setting("switchbarColourEvent", eSettingType.estNumber)
    m_colourMessage.colour = settings.setting("switchbarColourMessage", eSettingType.estNumber)
    m_colourAlert.colour = settings.setting("switchbarColourAlert", eSettingType.estNumber)
    m_colourHighlight.colour = settings.setting("switchbarColourHighlight", eSettingType.estNumber)
    
    getRealWindow(m_colourEvent).Move 425, 115, 20, 20
    getRealWindow(m_colourMessage).Move 425, 140, 20, 20
    getRealWindow(m_colourAlert).Move 425, 165, 20, 20
    getRealWindow(m_colourHighlight).Move 425, 190, 20, 20
    
    Set m_checkSbEventFlash = addCheckBox(Controls, "Flash icon", 450, 115, 100, 20)
    Set m_checkSbMessageFlash = addCheckBox(Controls, "Flash icon", 450, 140, 100, 20)
    Set m_checkSbAlertFlash = addCheckBox(Controls, "Flash icon", 450, 165, 100, 20)
    Set m_checkSbHighlightFlash = addCheckBox(Controls, "Flash icon", 450, 190, 100, 20)
    
    m_checkSbEventFlash.value = -settings.setting("switchbarFlashEvent", eSettingType.estBoolean)
    m_checkSbMessageFlash.value = -settings.setting("switchbarFlashMessage", eSettingType.estBoolean)
    m_checkSbAlertFlash.value = -settings.setting("switchbarFlashAlert", eSettingType.estBoolean)
    m_checkSbHighlightFlash.value = -settings.setting("switchbarFlashHighlight", eSettingType.estBoolean)
    
    m_labelManager.addLabel "Switchbar position:", ltNormal, 320, 220
    m_labelManager.addLabel "Switchbar rows:", ltNormal, 320, 270
    
    comboSwitchbarPosition.addItem "Top"
    comboSwitchbarPosition.addItem "Bottom"
    
    If StrComp(settings.setting("switchbarPosition", estString), "bottom", vbTextCompare) = 0 Then
        comboSwitchbarPosition.ListIndex = 1
    Else
        comboSwitchbarPosition.ListIndex = 0
    End If
    
    Dim count As Long
    
    For count = 1 To 10
        comboSwitchbarRows.addItem count
    Next count
    
    Dim index As Long
    
    index = settings.setting("switchbarRows", estNumber) - 1
    
    If index >= 0 And index < comboSwitchbarRows.ListCount Then
        comboSwitchbarRows.ListIndex = index
    Else
        comboSwitchbarRows.ListIndex = 0
    End If
    
    Set m_labelFontInfo = m_labelManager.addLabel("Current font: " & m_fontName & " size " & _
        CStr(m_fontSize), ltNormal, 320, 35)

    Set m_ctlEventColourEditor = createControl(Controls, "swiftIrc.ctlEventColourEditor", _
        "eventColourEditor")
    Set m_ctlColourPalette = createControl(Controls, "swiftIrc.ctlColourPalette", "colourPalette")
    
    getRealWindow(m_ctlEventColourEditor).Move 10, 100, 280, 155
    getRealWindow(m_ctlColourPalette).Move 10, 260, 200, 50
    
    m_ctlColourPalette.allowPaletteChange = True
    
    Set m_buttonNewTheme = addButton(Controls, "N&ew", 10, 35, 75, 20)
    Set m_buttonDeleteTheme = addButton(Controls, "&Delete", 90, 35, 75, 20)
    
    Set m_buttonChangeFont = addButton(Controls, "C&hange font", 360, 55, 100, 20)
    Set m_buttonDefaultFont = addButton(Controls, "&Restore default", 465, 55, 100, 20)
End Sub

Private Sub m_buttonChangeFont_clicked()
    Dim cf As tChooseFont
    
    cf.lStructSize = Len(cf)
    cf.hwndOwner = UserControl.hwnd
    cf.flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    cf.nSizeMin = 8
    cf.nSizeMax = 28
    
    Dim lf As LOGFONT
    Dim result As Long
    
    lf = m_fontmanager.fontStruct
    
    cf.lpLogFont = VarPtr(lf)
    
    result = ChooseFont(cf)
    
    If result = 0 Then
        Exit Sub
    End If
    
    m_fontName = StrConv(lf.lfFaceName, vbUnicode)
    m_fontSize = cf.iPointSize / 10
    
    Dim count As Long
    
    For count = 1 To Len(m_fontName)
        If Mid$(m_fontName, count, 1) = Chr$(0) Then
            m_fontName = Mid$(m_fontName, 1, count - 1)
        End If
    Next count
    
    If lf.lfWeight = FW_BOLD Then
        m_fontBold = True
    Else
        m_fontBold = False
    End If
    
    If lf.lfItalic <> 0 Then
        m_fontItalic = True
    Else
        m_fontItalic = False
    End If
    
    m_fontmanager.changeFont UserControl.hdc, m_fontName, m_fontSize, m_fontBold, m_fontItalic
    m_labelFontInfo.caption = "Current font: " & m_fontName & " size " & _
        CStr(m_fontSize)
    UserControl_Paint
End Sub

Private Sub m_buttonDefaultFont_clicked()
    m_fontName = getBestDefaultFont
    m_fontSize = 9
    m_fontBold = False
    m_fontItalic = False
    
    m_labelFontInfo.caption = "Current font: " & m_fontName & " size " & _
        m_fontSize
        
    UserControl_Paint
End Sub

Private Sub m_buttonDeleteTheme_clicked()
    If comboTheme.ListIndex = -1 Then
        Exit Sub
    End If

    If m_themes.count = 1 Then
        Exit Sub
    End If
    
    Dim index As Long
    
    index = comboTheme.ListIndex
    
    m_themes.Remove comboTheme.ListIndex + 1
    comboTheme.removeItem comboTheme.ListIndex
    
    If index > 0 Then
        comboTheme.ListIndex = index - 1
    Else
        comboTheme.ListIndex = 0
    End If

    Set m_currentTheme = m_themes.item(comboTheme.ListIndex + 1)
    loadTheme m_currentTheme
End Sub

Private Sub m_buttonNewTheme_clicked()
    Dim theme As CColourTheme
    Dim result As Variant
    Dim name As String
    
    result = requestInput("New theme", "Enter new theme name:", vbNullString, Me)
    
    If result = False Then
        Exit Sub
    End If
    
    name = result
    
    If LenB(name) = 0 Then
        Exit Sub
    End If
    
    If Not findTheme(name) Is Nothing Then
        MsgBox "A theme already exists by this name", vbCritical, "Could not create theme"
        Exit Sub
    End If
    
    Set theme = New CColourTheme
    
    m_currentTheme.copy theme
    
    theme.name = name
    m_themes.Add theme, LCase$(name)
    
    Set m_currentTheme = theme
    
    comboTheme.addItem theme.name
    comboTheme.ListIndex = comboTheme.ListCount - 1
End Sub

Private Function findTheme(name As String) As CColourTheme
    On Error Resume Next
    Set findTheme = m_themes.item(LCase$(name))
End Function

Private Sub m_colourMessage_colourChanged()
    m_sbTabMessage.foreColour = m_currentTheme.paletteEntry(m_colourMessage.colour)
    UserControl_Paint
End Sub

Private Sub m_colourHighlight_colourChanged()
    m_sbTabHighlight.foreColour = m_currentTheme.paletteEntry(m_colourHighlight.colour)
    UserControl_Paint
End Sub

Private Sub m_colourAlert_colourChanged()
    m_sbTabAlert.foreColour = m_currentTheme.paletteEntry(m_colourAlert.colour)
    UserControl_Paint
End Sub

Private Sub m_colourEvent_colourChanged()
    m_sbTabEvent.foreColour = m_currentTheme.paletteEntry(m_colourEvent.colour)
    UserControl_Paint
End Sub

Private Sub m_ctlColourPalette_colourChanged(index As Long, colour As Long)
    m_ctlEventColourEditor.paletteEntryUpdated index, colour
    m_currentTheme.paletteEntry(index) = colour
    
    m_colourEvent.setPalette m_currentTheme.getPalette
    m_colourMessage.setPalette m_currentTheme.getPalette
    m_colourAlert.setPalette m_currentTheme.getPalette
    m_colourHighlight.setPalette m_currentTheme.getPalette
End Sub

Private Sub m_ctlColourPalette_colourSelected(index As Long)
    m_ctlEventColourEditor.colourSelected index
    
    If m_ctlEventColourEditor.backgroundSelected Then
        m_currentTheme.backgroundColour = index
    Else
        m_currentTheme.eventColour(m_ctlEventColourEditor.selectedItem) = index
    End If
End Sub

Private Sub m_ctlEventColourEditor_itemClicked(colourIndex As Byte)
    m_ctlColourPalette.setFocus colourIndex
End Sub

Private Sub UserControl_Initialize()
    m_transTimer = 1
    
    SetTimer UserControl.hwnd, m_transTimer, 500, 0
    AttachMessage Me, UserControl.hwnd, WM_TIMER

    Set m_fontmanager = New CFontManager
    
    m_fontName = settings.fontName
    m_fontSize = settings.fontSize
    m_fontBold = settings.setting("fontBold", estBoolean)
    m_fontItalic = settings.setting("fontItalic", estBoolean)
    
    m_fontmanager.changeFont UserControl.hdc, m_fontName, m_fontSize, m_fontBold, m_fontItalic

    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    initControls
    
    Dim count As Long
    Dim index As Long
    Dim theme As CColourTheme
    
    colourThemes.copyThemes m_themes
    
    For Each theme In m_themes
        comboTheme.addItem theme.name
        
        count = count + 1
        
        If index = 0 Then
            If theme.name = colourThemes.currentTheme.name Then
                index = count
            End If
        End If
    Next theme
    
    Set m_currentTheme = m_themes.item(index)
    comboTheme.ListIndex = index - 1
    
    loadTheme m_currentTheme
End Sub

Private Sub loadTheme(theme As CColourTheme)
    colourThemes.currentSettingsTheme = theme

    m_ctlEventColourEditor.backgroundColour = theme.backgroundColour
    m_ctlEventColourEditor.setPalette theme.getPalette()
    m_ctlEventColourEditor.setEventColours theme.getEventColours()
    m_ctlColourPalette.setPalette theme.getPalette()
    
    m_colourEvent.setPalette theme.getPalette
    m_colourMessage.setPalette theme.getPalette
    m_colourAlert.setPalette theme.getPalette
    m_colourHighlight.setPalette theme.getPalette
End Sub

Public Sub saveSettings()
    colourThemes.clear
    
    Dim theme As CColourTheme
    
    For Each theme In m_themes
        colourThemes.addThemeIndirect theme
    Next theme
    
    colourThemes.currentTheme = m_currentTheme
    colourThemes.saveThemes
    
    eventColours.loadTheme colourThemes.currentTheme
    g_textViewBack = colourThemes.currentTheme.backgroundColour
    g_textViewFore = eventColours.normalText.colour
    g_textInputBack = colourThemes.currentTheme.backgroundColour
    g_textInputFore = eventColours.normalText.colour
    g_nicklistBack = colourThemes.currentTheme.backgroundColour
    g_nicklistFore = eventColours.normalText.colour
    g_channelListBack = colourThemes.currentTheme.backgroundColour
    g_channelListFore = eventColours.normalText.colour
    
    settings.setting("switchbarColourEvent", estNumber) = m_colourEvent.colour
    settings.setting("switchbarColourMessage", estNumber) = m_colourMessage.colour
    settings.setting("switchbarColourAlert", estNumber) = m_colourAlert.colour
    settings.setting("switchbarColourHighlight", estNumber) = m_colourHighlight.colour
    
    settings.setting("switchbarFlashEvent", estBoolean) = -m_checkSbEventFlash.value
    settings.setting("switchbarFlashMessage", estBoolean) = -m_checkSbMessageFlash.value
    settings.setting("switchbarFlashAlert", estBoolean) = -m_checkSbAlertFlash.value
    settings.setting("switchbarFlashHighlight", estBoolean) = -m_checkSbHighlightFlash.value
    
    If comboSwitchbarPosition.ListIndex = 0 Then
        settings.setting("switchbarPosition", estString) = "Top"
    Else
        settings.setting("switchbarPosition", estString) = "Bottom"
    End If
    
    settings.setting("switchbarRows", estNumber) = comboSwitchbarRows.ListIndex + 1
    
    settings.fontName = m_fontName
    settings.fontSize = m_fontSize
    settings.setting("fontBold", estBoolean) = m_fontBold
    settings.setting("fontItalic", estBoolean) = m_fontItalic
End Sub

Private Sub UserControl_Paint()
    Dim backBuffer As Long
    Dim backBitmap As Long
    
    Dim oldBitmap As Long
    Dim oldFont As Long
    
    backBuffer = CreateCompatibleDC(UserControl.hdc)
    backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
    
    oldBitmap = SelectObject(backBuffer, backBitmap)
    SetBkColor backBuffer, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    
    oldFont = SelectObject(backBuffer, GetCurrentObject(UserControl.hdc, OBJ_FONT))

    FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
        colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    
    m_labelManager.renderLabels backBuffer
    
    FrameRect backBuffer, makeRect(5, 305, 5, UserControl.ScaleHeight - 5), _
        colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
    FrameRect backBuffer, makeRect(315, UserControl.ScaleWidth - 5, 5, 80), _
        colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
    
    FrameRect backBuffer, makeRect(315, UserControl.ScaleWidth - 5, 85, UserControl.ScaleHeight - 5), _
        colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
    
    m_sbTabEvent.render backBuffer, 320, 115, 100, 20
    m_sbTabMessage.render backBuffer, 320, 140, 100, 20
    m_sbTabAlert.render backBuffer, 320, 165, 100, 20
    m_sbTabHighlight.render backBuffer, 320, 190, 100, 20
    
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        backBuffer, 0, 0, vbSrcCopy
    
    SelectObject backBuffer, oldBitmap
    SelectObject backBuffer, oldFont
    
    DeleteDC backBuffer
    DeleteObject backBitmap
End Sub

Private Sub UserControl_Resize()
    comboTheme.left = 10
    comboTheme.top = 65
    comboTheme.width = 280
    
    comboSwitchbarPosition.left = 320
    comboSwitchbarPosition.top = 240
    comboSwitchbarPosition.width = 75
    
    comboSwitchbarRows.left = 320
    comboSwitchbarRows.top = 290
    comboSwitchbarRows.width = 75
End Sub

Private Sub UserControl_Terminate()
    DetachMessage Me, UserControl.hwnd, WM_TIMER
End Sub
