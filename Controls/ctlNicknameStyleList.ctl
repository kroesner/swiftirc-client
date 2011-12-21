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
    Set m_client = newValue
End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_DRAWITEM
            ISubclass_MsgResponse = emrConsume
        Case Else
            ISubclass_MsgResponse = emrConsume
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
        Case WM_DRAWITEM
            Dim item As DRAWITEMSTRUCT
            
            CopyMemory item, ByVal lParam, Len(item)
            ISubclass_WindowProc = drawItem(item)
        Case WM_MEASUREITEM
            Dim measureItem As MEASUREITEMSTRUCT
            Dim tm As TEXTMETRIC
        
            UserControl.fontName = settings.fontName
            UserControl.fontSize = settings.fontSize
            GetTextMetrics UserControl.hdc, tm
            m_itemHeight = tm.tmHeight
            
            CopyMemory measureItem, ByVal lParam, Len(measureItem)
            measureItem.itemHeight = m_itemHeight
            CopyMemory ByVal lParam, measureItem, Len(measureItem)
        
            ISubclass_WindowProc = True
        Case WM_CTLCOLORLISTBOX
            If m_fillbrush <> 0 Then
                DeleteObject m_fillbrush
            End If
        
            m_fillbrush = CreateSolidBrush(getPaletteEntry(g_nicklistBack))
            
            ISubclass_WindowProc = m_fillbrush
    End Select
End Function

Private Function drawItem(item As DRAWITEMSTRUCT) As Long
    If item.itemAction = ODA_FOCUS Then
        If item.itemState And ODS_FOCUS Then
            DrawFocusRect item.hdc, item.rcItem
            drawItem = 1
            Exit Function
        End If
    End If
    
    If item.itemID = -1 Then
        Exit Function
    End If
    
    Dim icon As CImage
    Dim fillBrush As Long
    Dim textColour As Long
    Dim text As String
    
    Select Case item.itemID
        Case m_itemOps
            textColour = getSettingsPaletteEntry(m_styleOps.foreColour)
            Set icon = m_styleOps.image
            text = "Ops"
        Case m_itemHalfOps
            textColour = getSettingsPaletteEntry(m_styleHalfOps.foreColour)
            Set icon = m_styleHalfOps.image
            text = "Halfops"
        Case m_itemVoices
            textColour = getSettingsPaletteEntry(m_styleVoices.foreColour)
            Set icon = m_styleVoices.image
            text = "Voices"
        Case m_itemMe
            textColour = getSettingsPaletteEntry(m_styleMe.foreColour)
            Set icon = m_styleMe.image
            text = "Me"
        Case m_itemNormal
            textColour = getSettingsPaletteEntry(m_styleNormal.foreColour)
            Set icon = m_styleNormal.image
            text = "Normal"
    End Select

    If item.itemState And ODS_SELECTED Then
        SetTextColor item.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
        SetBkColor item.hdc, GetSysColor(COLOR_HIGHLIGHT)
        FillRect item.hdc, item.rcItem, GetSysColorBrush(COLOR_HIGHLIGHT)
    Else
        fillBrush = CreateSolidBrush(getSettingsPaletteEntry(g_nicklistBack))
        SetBkColor item.hdc, getSettingsPaletteEntry(g_nicklistBack)
        FillRect item.hdc, item.rcItem, fillBrush
        DeleteObject fillBrush
        SetTextColor item.hdc, textColour
    End If
    
    Dim textRect As RECT
    
    textRect = item.rcItem
    
    'SelectObject item.hdc, GetCurrentObject(UserControl.hdc, OBJ_FONT)
    
    Dim newFont As Long
    Dim oldFont As Long
    
    If m_checkBoldNicks.value Then
        newFont = m_client.fontManager.getFont(True, False, False)
    Else
        newFont = m_client.fontManager.getDefaultFont
    End If
    
    oldFont = SelectObject(item.hdc, newFont)
    
    If Not icon Is Nothing And m_drawIcons Then
        icon.draw item.hdc, item.rcItem.left, item.rcItem.top, m_itemHeight, m_itemHeight
        textRect.left = textRect.left + m_itemHeight + 1
        swiftTextOut item.hdc, textRect.left, textRect.top, ETO_OPAQUE, VarPtr(textRect), _
            text
    Else
        swiftTextOut item.hdc, textRect.left + 5, textRect.top, ETO_OPAQUE, VarPtr(textRect), _
            text
    End If
    
    SelectObject item.hdc, oldFont
End Function

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
    m_labelManager.addLabel "Nickname appearance", ltHeading, 5, 5
    
    Set m_checkNickIcons = addCheckBox(Controls, "Nickname icons", 165, 30, 130, 20)
    Set m_checkColourMyNick = addCheckBox(Controls, "Change colour of my nickname", 165, 50, 130, 30)
    Set m_checkBoldNicks = addCheckBox(Controls, "Bold nicknames", 165, 80, 130, 20)
    
    m_checkNickIcons.value = -settings.setting("nicknameIcons", estBoolean)
    m_checkColourMyNick.value = -settings.setting("colourMyNick", estBoolean)
    m_checkBoldNicks.value = -settings.setting("boldNicks", estBoolean)
    
    Set m_colourPalette = createControl(Controls, "swiftIrc.ctlColourPalette", "palette")
    m_colourPalette.setPalette colourThemes.currentSettingsTheme.getPalette
    
    m_listbox = CreateWindowEx(0, "LISTBOX", 0&, WS_CHILD Or _
        WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL Or LBS_HASSTRINGS _
        Or LBS_OWNERDRAWFIXED Or LBS_NOINTEGRALHEIGHT, _
        0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        UserControl.hwnd, 0&, App.hInstance, 0&)
    
    m_itemOps = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Ops")
    m_itemHalfOps = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Halfops")
    m_itemVoices = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Voices")
    m_itemMe = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Me")
    m_itemNormal = SendMessage(m_listbox, LB_ADDSTRING, 0, ByVal "Normal")
    
    updateColours Controls
End Sub

Private Sub m_checkBoldNicks_Click()
    RedrawWindow m_listbox, ByVal 0, 0, RDW_INVALIDATE
End Sub

Private Sub m_checkNickIcons_Click()
    m_drawIcons = -m_checkNickIcons.value
    RedrawWindow m_listbox, ByVal 0, ByVal 0, RDW_INVALIDATE
End Sub

Private Sub m_colourPalette_colourSelected(index As Long)
    Dim selectedItem As Long
    
    selectedItem = SendMessage(m_listbox, LB_GETCURSEL, 0, ByVal 0)
    
    Select Case selectedItem
        Case m_itemOps
            m_styleOps.foreColour = index
        Case m_itemHalfOps
            m_styleHalfOps.foreColour = index
        Case m_itemVoices
            m_styleVoices.foreColour = index
        Case m_itemMe
            m_styleMe.foreColour = index
        Case m_itemNormal
            m_styleNormal.foreColour = index
    End Select
    
    RedrawWindow m_listbox, ByVal 0, 0, RDW_INVALIDATE
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    
    Set m_styleOps = prefixStyles.item(1).style.copy
    Set m_styleHalfOps = prefixStyles.item(2).style.copy
    Set m_styleVoices = prefixStyles.item(3).style.copy
    Set m_styleMe = styleMe.copy
    Set m_styleNormal = styleNormal.copy
    
    initMessages
    initControls
End Sub

Private Sub UserControl_Paint()
    Dim backBuffer As Long
    Dim backBitmap As Long
    Dim oldBitmap As Long
    Dim oldFont As Long
    
    backBuffer = CreateCompatibleDC(UserControl.hdc)
    backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)

    oldBitmap = SelectObject(backBuffer, backBitmap)
    oldFont = SelectObject(backBuffer, GetCurrentObject(UserControl.hdc, OBJ_FONT))

    SetBkColor backBuffer, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)

    FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
        colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    
    m_labelManager.renderLabels backBuffer
    
    FrameRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
        colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
        
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, _
        0, 0, vbSrcCopy
    
    SelectObject backBuffer, oldBitmap
    SelectObject backBuffer, oldFont
    
    DeleteDC backBuffer
    DeleteObject backBitmap
End Sub

Friend Sub saveSettings()
    settings.setting("nicknameIcons", estBoolean) = -m_checkNickIcons.value
    settings.setting("colourMyNick", estBoolean) = -m_checkColourMyNick.value
    settings.setting("boldNicks", estBoolean) = -m_checkBoldNicks.value
    
    settings.setting("nickColourOps", estNumber) = m_styleOps.foreColour
    settings.setting("nickColourHalfOps", estNumber) = m_styleHalfOps.foreColour
    settings.setting("nickColourVoices", estNumber) = m_styleVoices.foreColour
    settings.setting("nickColourMe", estNumber) = m_styleMe.foreColour
    settings.setting("nickColourNormal", estNumber) = m_styleNormal.foreColour
    
    prefixStyles.item(1).style.foreColour = m_styleOps.foreColour
    prefixStyles.item(2).style.foreColour = m_styleHalfOps.foreColour
    prefixStyles.item(3).style.foreColour = m_styleVoices.foreColour
    
    styleMe.foreColour = m_styleMe.foreColour
    styleMeOp.foreColour = m_styleMe.foreColour
    styleMeHalfop.foreColour = m_styleMe.foreColour
    styleMeVoice.foreColour = m_styleMe.foreColour
    
    styleNormal.foreColour = m_styleNormal.foreColour
End Sub

Private Sub initMessages()
    AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
    AttachMessage Me, UserControl.hwnd, WM_MEASUREITEM
    AttachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub deInitMessages()
    DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
    DetachMessage Me, UserControl.hwnd, WM_MEASUREITEM
    DetachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub UserControl_Resize()
    MoveWindow m_listbox, 10, 30, 150, 75, 1
    getRealWindow(m_colourPalette).Move 5, 110, 200, 60
End Sub

Private Sub UserControl_Terminate()
    deInitMessages
End Sub
