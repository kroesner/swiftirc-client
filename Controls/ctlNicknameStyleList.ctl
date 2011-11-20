VERSION 5.00
Begin VB.UserControl ctlNicknameStyleList 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
End
Attribute VB_Name = "ctlNicknameStyleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements ISubclass

Private m_realWindow As VBControlExtender
Private m_labelManager As New CLabelManager
Private m_listbox As Long

Private m_client As SwiftIrcClient

Private WithEvents m_colourPalette As ctlColourPalette
Attribute m_colourPalette.VB_VarHelpID = -1
Private WithEvents m_checkNickIcons As VB.CheckBox
Attribute m_checkNickIcons.VB_VarHelpID = -1
Private WithEvents m_checkColourMyNick As VB.CheckBox
Attribute m_checkColourMyNick.VB_VarHelpID = -1
Private WithEvents m_checkBoldNicks As VB.CheckBox
Attribute m_checkBoldNicks.VB_VarHelpID = -1

Private m_itemOps As Long
Private m_itemHalfOps As Long
Private m_itemVoices As Long
Private m_itemMe As Long
Private m_itemNormal As Long

Private m_styleOps As New CUserStyle
Private m_styleHalfOps As New CUserStyle
Private m_styleVoices As New CUserStyle
Private m_styleMe As New CUserStyle
Private m_styleNormal As New CUserStyle

Private m_drawIcons As Boolean

Private m_fillbrush As Long
Private m_itemHeight As Long

Public Property Let client(newValue As SwiftIrcClient)
10        Set m_client = newValue
End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        Select Case CurrentMessage
              Case WM_DRAWITEM
20                ISubclass_MsgResponse = emrConsume
30            Case Else
40                ISubclass_MsgResponse = emrConsume
50        End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
10        Select Case iMsg
              Case WM_DRAWITEM
                  Dim item As DRAWITEMSTRUCT
                  
20                CopyMemory item, ByVal lParam, Len(item)
30                ISubclass_WindowProc = drawItem(item)
40            Case WM_MEASUREITEM
                  Dim measureItem As MEASUREITEMSTRUCT
                  Dim tm As TEXTMETRIC
              
50                UserControl.fontName = settings.fontName
60                UserControl.fontSize = settings.fontSize
70                GetTextMetrics UserControl.hdc, tm
80                m_itemHeight = tm.tmHeight
                  
90                CopyMemory measureItem, ByVal lParam, Len(measureItem)
100               measureItem.itemHeight = m_itemHeight
110               CopyMemory ByVal lParam, measureItem, Len(measureItem)
              
120               ISubclass_WindowProc = True
130           Case WM_CTLCOLORLISTBOX
140               If m_fillbrush <> 0 Then
150                   DeleteObject m_fillbrush
160               End If
              
170               m_fillbrush = CreateSolidBrush(getPaletteEntry(g_nicklistBack))
                  
180               ISubclass_WindowProc = m_fillbrush
190       End Select
End Function

Private Function drawItem(item As DRAWITEMSTRUCT) As Long
10        If item.itemAction = ODA_FOCUS Then
20            If item.itemState And ODS_FOCUS Then
30                DrawFocusRect item.hdc, item.rcItem
40                drawItem = 1
50                Exit Function
60            End If
70        End If
          
80        If item.itemID = -1 Then
90            Exit Function
100       End If
          
          Dim icon As CImage
          Dim fillBrush As Long
          Dim textColour As Long
          Dim text As String
          
110       Select Case item.itemID
              Case m_itemOps
120               textColour = getSettingsPaletteEntry(m_styleOps.foreColour)
130               Set icon = m_styleOps.image
140               text = "Ops"
150           Case m_itemHalfOps
160               textColour = getSettingsPaletteEntry(m_styleHalfOps.foreColour)
170               Set icon = m_styleHalfOps.image
180               text = "Halfops"
190           Case m_itemVoices
200               textColour = getSettingsPaletteEntry(m_styleVoices.foreColour)
210               Set icon = m_styleVoices.image
220               text = "Voices"
230           Case m_itemMe
240               textColour = getSettingsPaletteEntry(m_styleMe.foreColour)
250               Set icon = m_styleMe.image
260               text = "Me"
270           Case m_itemNormal
280               textColour = getSettingsPaletteEntry(m_styleNormal.foreColour)
290               Set icon = m_styleNormal.image
300               text = "Normal"
310       End Select

320       If item.itemState And ODS_SELECTED Then
330           SetTextColor item.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
340           SetBkColor item.hdc, GetSysColor(COLOR_HIGHLIGHT)
350           FillRect item.hdc, item.rcItem, GetSysColorBrush(COLOR_HIGHLIGHT)
360       Else
370           fillBrush = CreateSolidBrush(getSettingsPaletteEntry(g_nicklistBack))
380           SetBkColor item.hdc, getSettingsPaletteEntry(g_nicklistBack)
390           FillRect item.hdc, item.rcItem, fillBrush
400           DeleteObject fillBrush
410           SetTextColor item.hdc, textColour
420       End If
          
          Dim textRect As RECT
          
430       textRect = item.rcItem
          
          'SelectObject item.hdc, GetCurrentObject(UserControl.hdc, OBJ_FONT)
          
          Dim newFont As Long
          Dim oldFont As Long
          
440       If m_checkBoldNicks.value Then
450           newFont = m_client.fontManager.getFont(True, False, False)
460       Else
470           newFont = m_client.fontManager.getDefaultFont
480       End If
          
490       oldFont = SelectObject(item.hdc, newFont)
          
500       If Not icon Is Nothing And m_drawIcons Then
510           icon.draw item.hdc, item.rcItem.left, item.rcItem.top, m_itemHeight, m_itemHeight
520           textRect.left = textRect.left + m_itemHeight + 1
530           swiftTextOut item.hdc, textRect.left, textRect.top, ETO_OPAQUE, VarPtr(textRect), _
                  text
540       Else
550           swiftTextOut item.hdc, textRect.left + 5, textRect.top, ETO_OPAQUE, VarPtr(textRect), _
                  text
560       End If
          
570       SelectObject item.hdc, oldFont
End Function

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
10        m_labelManager.addLabel "Nickname appearance", ltHeading, 5, 5
          
20        Set m_checkNickIcons = addCheckBox(Controls, "Nickname icons", 165, 30, 130, 20)
30        Set m_checkColourMyNick = addCheckBox(Controls, "Change colour of my nickname", 165, 50, 130, 30)
40        Set m_checkBoldNicks = addCheckBox(Controls, "Bold nicknames", 165, 80, 130, 20)
          
50        m_checkNickIcons.value = -settings.setting("nicknameIcons", estBoolean)
60        m_checkColourMyNick.value = -settings.setting("colourMyNick", estBoolean)
70        m_checkBoldNicks.value = -settings.setting("boldNicks", estBoolean)
          
80        Set m_colourPalette = createControl(Controls, "swiftIrc.ctlColourPalette", "palette")
90        m_colourPalette.setPalette colourThemes.currentSettingsTheme.getPalette
          
100       m_listbox = CreateWindowEx(0, "LISTBOX", 0&, WS_CHILD Or _
              WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL Or LBS_HASSTRINGS _
              Or LBS_OWNERDRAWFIXED Or LBS_NOINTEGRALHEIGHT, _
              0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
              UserControl.hwnd, 0&, App.hInstance, 0&)
          
110       m_itemOps = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Ops")
120       m_itemHalfOps = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Halfops")
130       m_itemVoices = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Voices")
140       m_itemMe = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Me")
150       m_itemNormal = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Normal")
          
160       updateColours Controls
End Sub

Private Sub m_checkBoldNicks_Click()
10        RedrawWindow m_listbox, ByVal 0, 0, RDW_INVALIDATE
End Sub

Private Sub m_checkNickIcons_Click()
10        m_drawIcons = -m_checkNickIcons.value
20        RedrawWindow m_listbox, ByVal 0, ByVal 0, RDW_INVALIDATE
End Sub

Private Sub m_colourPalette_colourSelected(index As Long)
          Dim selectedItem As Long
          
10        selectedItem = SendMessage(m_listbox, LB_GETCURSEL, 0, ByVal 0)
          
20        Select Case selectedItem
              Case m_itemOps
30                m_styleOps.foreColour = index
40            Case m_itemHalfOps
50                m_styleHalfOps.foreColour = index
60            Case m_itemVoices
70                m_styleVoices.foreColour = index
80            Case m_itemMe
90                m_styleMe.foreColour = index
100           Case m_itemNormal
110               m_styleNormal.foreColour = index
120       End Select
          
130       RedrawWindow m_listbox, ByVal 0, 0, RDW_INVALIDATE
End Sub

Private Sub UserControl_Initialize()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
          
20        Set m_styleOps = prefixStyles.item(1).style.copy
30        Set m_styleHalfOps = prefixStyles.item(2).style.copy
40        Set m_styleVoices = prefixStyles.item(3).style.copy
50        Set m_styleMe = styleMe.copy
60        Set m_styleNormal = styleNormal.copy
          
70        initMessages
80        initControls
End Sub

Private Sub UserControl_Paint()
          Dim backBuffer As Long
          Dim backBitmap As Long
          Dim oldBitmap As Long
          Dim oldFont As Long
          
10        backBuffer = CreateCompatibleDC(UserControl.hdc)
20        backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)

30        oldBitmap = SelectObject(backBuffer, backBitmap)
40        oldFont = SelectObject(backBuffer, GetCurrentObject(UserControl.hdc, OBJ_FONT))

50        SetBkColor backBuffer, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)

60        FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
          
70        m_labelManager.renderLabels backBuffer
          
80        FrameRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
              
90        BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, _
              0, 0, vbSrcCopy
          
100       SelectObject backBuffer, oldBitmap
110       SelectObject backBuffer, oldFont
          
120       DeleteDC backBuffer
130       DeleteObject backBitmap
End Sub

Friend Sub saveSettings()
10        settings.setting("nicknameIcons", estBoolean) = -m_checkNickIcons.value
20        settings.setting("colourMyNick", estBoolean) = -m_checkColourMyNick.value
30        settings.setting("boldNicks", estBoolean) = -m_checkBoldNicks.value
          
40        settings.setting("nickColourOps", estNumber) = m_styleOps.foreColour
50        settings.setting("nickColourHalfOps", estNumber) = m_styleHalfOps.foreColour
60        settings.setting("nickColourVoices", estNumber) = m_styleVoices.foreColour
70        settings.setting("nickColourMe", estNumber) = m_styleMe.foreColour
80        settings.setting("nickColourNormal", estNumber) = m_styleNormal.foreColour
          
90        prefixStyles.item(1).style.foreColour = m_styleOps.foreColour
100       prefixStyles.item(2).style.foreColour = m_styleHalfOps.foreColour
110       prefixStyles.item(3).style.foreColour = m_styleVoices.foreColour
          
120       styleMe.foreColour = m_styleMe.foreColour
130       styleMeOp.foreColour = m_styleMe.foreColour
140       styleMeHalfop.foreColour = m_styleMe.foreColour
150       styleMeVoice.foreColour = m_styleMe.foreColour
          
160       styleNormal.foreColour = m_styleNormal.foreColour
End Sub

Private Sub initMessages()
10        AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
20        AttachMessage Me, UserControl.hwnd, WM_MEASUREITEM
30        AttachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub deInitMessages()
10        DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
20        DetachMessage Me, UserControl.hwnd, WM_MEASUREITEM
30        DetachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub UserControl_Resize()
10        MoveWindow m_listbox, 10, 30, 150, 75, 1
20        getRealWindow(m_colourPalette).Move 5, 110, 200, 60
End Sub

Private Sub UserControl_Terminate()
10        deInitMessages
End Sub
