VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_left As Integer
Private m_right As Integer
Private m_top As Integer
Private m_bottom As Integer
Private m_visible As Boolean

Private m_region As Long
Private m_caption As String
Private m_window As IWindow

Private m_foreColour As Long
Private m_icon As CImage
Private m_trans As Boolean

Private m_state As eSwitchbarTabState
Private m_activityState As eTabActivityState
Private m_order As eSwitchbarOrder

Private Const TAB_MARGIN_TOP = 5
Private Const TAB_MARGIN_BOTTOM = 5
Private Const TAB_MARGIN_LEFT = 5
Private Const TAB_MARGIN_RIGHT = 5
Private Const TAB_ICON_PADDING = 5

Public Property Get order() As eSwitchbarOrder
    order = m_order
End Property

Public Property Let order(newValue As eSwitchbarOrder)
    m_order = newValue
End Property

Public Property Get trans() As Boolean
    trans = m_trans
End Property

Public Property Let trans(newValue As Boolean)
    m_trans = newValue
End Property

Public Property Get left() As Integer
    left = m_left
End Property

Public Property Let left(newValue As Integer)
    m_left = newValue
End Property

Public Property Get right() As Integer
    right = m_right
End Property

Public Property Let right(newValue As Integer)
    m_right = newValue
End Property

Public Property Get top() As Integer
    top = m_top
End Property

Public Property Let top(newValue As Integer)
    m_top = newValue
End Property

Public Property Get bottom() As Integer
    bottom = m_bottom
End Property

Public Property Let bottom(newValue As Integer)
    m_bottom = newValue
End Property

Public Property Get visible() As Boolean
    visible = m_visible
End Property

Public Property Let visible(newValue As Boolean)
    m_visible = newValue
End Property

Public Property Get window() As IWindow
    Set window = m_window
End Property

Public Property Let window(newValue As IWindow)
    Set m_window = newValue
End Property

Public Property Get caption() As String
    caption = m_caption
End Property

Public Property Let caption(newValue As String)
    m_caption = newValue
End Property

Public Property Get icon() As CImage
    Set icon = m_icon
End Property

Public Property Let icon(newValue As CImage)
    Set m_icon = newValue
End Property

Public Property Get foreColour() As Long
    foreColour = m_foreColour
End Property

Public Property Let foreColour(newValue As Long)
    m_foreColour = newValue
End Property

Public Property Get activityState() As eTabActivityState
    activityState = m_activityState
End Property

Public Property Let activityState(newValue As eTabActivityState)
    m_activityState = newValue
End Property

Public Property Get state() As eSwitchbarTabState
    state = m_state
End Property

Public Property Let state(newValue As eSwitchbarTabState)
    m_state = newValue
End Property

Public Sub updateRegion()
    If m_region <> 0 Then
        DeleteObject m_region
    End If
    
    m_region = CreateRoundRectRgn(m_left, m_top, m_right, m_bottom, 7, 7)
End Sub

Public Function mouseOverTab(x As Long, y As Long) As Boolean
    If PtInRegion(m_region, x, y) <> 0 Then
        mouseOverTab = True
    End If
End Function

Public Sub render(hdc As Long, x As Long, y As Long, width As Long, height As Long)
    Dim tabRect As RECT
    
    tabRect.left = x
    tabRect.top = y
    tabRect.right = x + width
    tabRect.bottom = y + height
    
    Dim oldBrush As Long
    Dim oldPen As Long
    Dim backBrush As Long
    
    If g_initialized Then
        If m_state = stsSelected Then
            backBrush = colourManager.getBrush(SWIFTCOLOUR_WINDOW)
            oldBrush = SelectObject(hdc, backBrush)
            oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_BORDER))
        Else
            backBrush = colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
            oldBrush = SelectObject(hdc, backBrush)
            oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_BORDER))
        End If
    End If
    
    RoundRect hdc, x, y, x + width, y + height, 7, 7
    SelectObject hdc, oldBrush
    SelectObject hdc, oldPen
    
    If g_initialized Then
        If m_foreColour <> -1 Then
            SetTextColor hdc, m_foreColour
        ElseIf m_state = stsNormal Then
            If m_activityState = tasNormal Then
                SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
            ElseIf m_activityState = tasEvent Then
                SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourEvent", eSettingType.estNumber))
            ElseIf m_activityState = tasMessage Then
                SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourMessage", eSettingType.estNumber))
            ElseIf m_activityState = tasAlert Then
                SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourAlert", eSettingType.estNumber))
            ElseIf m_activityState = tasHighlight Then
                SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourHighlight", eSettingType.estNumber))
            End If
        ElseIf m_state = stsSelected Then
            SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
        ElseIf m_state = stsMouseOver Then
            SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
        Else
            SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
        End If
    End If
    
    SetBkMode hdc, TRANSPARENT
    
    Dim textRect As RECT
    
    textRect = tabRect
    textRect.left = textRect.left + TAB_MARGIN_LEFT
    textRect.right = textRect.right - TAB_MARGIN_RIGHT
    
    If Not m_icon Is Nothing Then
        Dim iconSize As Long
        Dim iconY As Long
        
        iconSize = 16
        
        If iconSize Mod 2 Then
            iconSize = iconSize - 1
        End If
        
        iconY = tabRect.top + ((tabRect.bottom - tabRect.top) / 2) - (iconSize / 2)
        
        If trans Then
            m_icon.drawAlpha hdc, tabRect.left + TAB_MARGIN_LEFT, iconY, iconSize, iconSize, backBrush
        Else
            m_icon.draw hdc, tabRect.left + TAB_MARGIN_LEFT, iconY, iconSize, iconSize
        End If
        textRect.left = textRect.left + iconSize + TAB_ICON_PADDING
    End If
    
    swiftDrawText hdc, m_caption, VarPtr(textRect), DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS
End Sub

Private Sub Class_Initialize()
    m_foreColour = -1
End Sub
