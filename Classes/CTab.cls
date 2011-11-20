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
10        order = m_order
End Property

Public Property Let order(newValue As eSwitchbarOrder)
10        m_order = newValue
End Property

Public Property Get trans() As Boolean
10        trans = m_trans
End Property

Public Property Let trans(newValue As Boolean)
10        m_trans = newValue
End Property

Public Property Get left() As Integer
10        left = m_left
End Property

Public Property Let left(newValue As Integer)
10        m_left = newValue
End Property

Public Property Get right() As Integer
10        right = m_right
End Property

Public Property Let right(newValue As Integer)
10        m_right = newValue
End Property

Public Property Get top() As Integer
10        top = m_top
End Property

Public Property Let top(newValue As Integer)
10        m_top = newValue
End Property

Public Property Get bottom() As Integer
10        bottom = m_bottom
End Property

Public Property Let bottom(newValue As Integer)
10        m_bottom = newValue
End Property

Public Property Get visible() As Boolean
10        visible = m_visible
End Property

Public Property Let visible(newValue As Boolean)
10        m_visible = newValue
End Property

Public Property Get window() As IWindow
10        Set window = m_window
End Property

Public Property Let window(newValue As IWindow)
10        Set m_window = newValue
End Property

Public Property Get caption() As String
10        caption = m_caption
End Property

Public Property Let caption(newValue As String)
10        m_caption = newValue
End Property

Public Property Get icon() As CImage
10        Set icon = m_icon
End Property

Public Property Let icon(newValue As CImage)
10        Set m_icon = newValue
End Property

Public Property Get foreColour() As Long
10        foreColour = m_foreColour
End Property

Public Property Let foreColour(newValue As Long)
10        m_foreColour = newValue
End Property

Public Property Get activityState() As eTabActivityState
10        activityState = m_activityState
End Property

Public Property Let activityState(newValue As eTabActivityState)
10        m_activityState = newValue
End Property

Public Property Get state() As eSwitchbarTabState
10        state = m_state
End Property

Public Property Let state(newValue As eSwitchbarTabState)
10        m_state = newValue
End Property

Public Sub updateRegion()
10        If m_region <> 0 Then
20            DeleteObject m_region
30        End If
          
40        m_region = CreateRoundRectRgn(m_left, m_top, m_right, m_bottom, 7, 7)
End Sub

Public Function mouseOverTab(x As Long, y As Long) As Boolean
10        If PtInRegion(m_region, x, y) <> 0 Then
20            mouseOverTab = True
30        End If
End Function

Public Sub render(hdc As Long, x As Long, y As Long, width As Long, height As Long)
          Dim tabRect As RECT
          
10        tabRect.left = x
20        tabRect.top = y
30        tabRect.right = x + width
40        tabRect.bottom = y + height
          
          Dim oldBrush As Long
          Dim oldPen As Long
          Dim backBrush As Long
          
50        If g_initialized Then
60            If m_state = stsSelected Then
70                backBrush = colourManager.getBrush(SWIFTCOLOUR_WINDOW)
80                oldBrush = SelectObject(hdc, backBrush)
90                oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_BORDER))
100           Else
110               backBrush = colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
120               oldBrush = SelectObject(hdc, backBrush)
130               oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_BORDER))
140           End If
150       End If
          
160       RoundRect hdc, x, y, x + width, y + height, 7, 7
170       SelectObject hdc, oldBrush
180       SelectObject hdc, oldPen
          
190       If g_initialized Then
200           If m_foreColour <> -1 Then
210               SetTextColor hdc, m_foreColour
220           ElseIf m_state = stsNormal Then
230               If m_activityState = tasNormal Then
240                   SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
250               ElseIf m_activityState = tasEvent Then
260                   SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourEvent", eSettingType.estNumber))
270               ElseIf m_activityState = tasMessage Then
280                   SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourMessage", eSettingType.estNumber))
290               ElseIf m_activityState = tasAlert Then
300                   SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourAlert", eSettingType.estNumber))
310               ElseIf m_activityState = tasHighlight Then
320                   SetTextColor hdc, getPaletteEntry(settings.setting("switchbarColourHighlight", eSettingType.estNumber))
330               End If
340           ElseIf m_state = stsSelected Then
350               SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
360           ElseIf m_state = stsMouseOver Then
370               SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
380           Else
390               SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
400           End If
410       End If
          
420       SetBkMode hdc, TRANSPARENT
          
          Dim textRect As RECT
          
430       textRect = tabRect
440       textRect.left = textRect.left + TAB_MARGIN_LEFT
450       textRect.right = textRect.right - TAB_MARGIN_RIGHT
          
460       If Not m_icon Is Nothing Then
              Dim iconSize As Long
              Dim iconY As Long
              
470           iconSize = 16
              
480           If iconSize Mod 2 Then
490               iconSize = iconSize - 1
500           End If
              
510           iconY = tabRect.top + ((tabRect.bottom - tabRect.top) / 2) - (iconSize / 2)
              
520           If trans Then
530               m_icon.drawAlpha hdc, tabRect.left + TAB_MARGIN_LEFT, iconY, iconSize, iconSize, backBrush
540           Else
550               m_icon.draw hdc, tabRect.left + TAB_MARGIN_LEFT, iconY, iconSize, iconSize
560           End If
570           textRect.left = textRect.left + iconSize + TAB_ICON_PADDING
580       End If
          
590       swiftDrawText hdc, m_caption, VarPtr(textRect), DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX _
              Or DT_END_ELLIPSIS
End Sub

Private Sub Class_Initialize()
10        m_foreColour = -1
End Sub
