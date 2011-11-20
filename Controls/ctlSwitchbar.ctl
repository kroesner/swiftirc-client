VERSION 5.00
Begin VB.UserControl ctlSwitchbar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Menu mnuTab 
      Caption         =   "mnuTab"
      Begin VB.Menu mnuTabSwitch 
         Caption         =   "Switch to"
      End
      Begin VB.Menu mnuTabClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "ctlSwitchbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IColourUser
Implements ISubclass

Public Enum eSwitchbarOrder
    sboStatus
    sboChannel
    sboQuery
    sboChannelList
    sboGeneric
End Enum

Private m_realWindow As VBControlExtender

Private m_position As eSwitchbarPosition
Private m_resizing As Boolean
Private m_moving As Boolean

Private m_timerFlash As Long
Private m_trans As Boolean

Private Const MARGIN_LEFT = 5
Private Const MARGIN_RIGHT = 5
Private Const MARGIN_TOP = 2
Private Const MARGIN_BOTTOM = 2

Private Const TAB_MARGIN_TOP = 2
Private Const TAB_MARGIN_BOTTOM = 2
Private Const TAB_MARGIN_LEFT = 3
Private Const TAB_MARGIN_RIGHT = 3

Private Const TAB_SPACING_X = 3
Private Const TAB_SPACING_Y = 3

Private MIN_TAB_WIDTH As Integer
Private Const MAX_TAB_WIDTH = 125

Private m_tabs As New cArrayList

Private m_tabHeight As Long
Private m_tabWidth As Long
Private m_rows As Long
Private m_tabsPerRow As Long

Private m_overTab As CTab
Private m_selectedTab As CTab
Private m_contextTab As CTab 'Tab associated with right click menu

Private m_fontManager As CFontManager

Public Event changeHeight(newHeight As Long)
Public Event moveRequest(x As Single, y As Single)
Public Event tabSelected(selectedTab As CTab)
Public Event closeRequest(aTab As CTab)

Public Property Get rows() As Long
10        rows = m_rows
End Property

Public Property Let rows(newValue As Long)
10        updateRowCount newValue
End Property

Private Sub IColourUser_coloursUpdated()
10        reDraw
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
10        Select Case iMsg
              Case WM_TIMER
20                If wParam = m_timerFlash Then
                      Dim count As Long
                      
30                    m_trans = Not m_trans
                      
40                    For count = 1 To m_tabs.count
50                        If m_tabs.item(count).activityState = eTabActivityState.tasEvent Then
60                            If settings.setting("switchbarFlashEvent", estBoolean) Then
70                                updateTabFlash m_tabs.item(count), m_trans
80                            End If
90                        ElseIf m_tabs.item(count).activityState = eTabActivityState.tasMessage Then
100                           If settings.setting("switchbarFlashMessage", estBoolean) Then
110                               updateTabFlash m_tabs.item(count), m_trans
120                           End If
130                       ElseIf m_tabs.item(count).activityState = eTabActivityState.tasAlert Then
140                           If settings.setting("switchbarFlashAlert", estBoolean) Then
150                               updateTabFlash m_tabs.item(count), m_trans
160                           End If
170                       ElseIf m_tabs.item(count).activityState = eTabActivityState.tasHighlight Then
180                           If settings.setting("switchbarFlashHighlight", estBoolean) Then
190                               updateTabFlash m_tabs.item(count), m_trans
200                           End If
210                       End If
220                   Next count
230               End If
240       End Select
End Function

Private Sub updateTabFlash(tabItem As CTab, trans As Boolean)
10        tabItem.trans = trans
20        redrawTab tabItem
End Sub

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Public Property Get position() As eSwitchbarPosition
10        position = m_position
End Property

Public Property Let position(newValue As eSwitchbarPosition)
10        m_position = newValue
          
20        If m_position = eSwitchbarPosition.sbpTop Then
30            settings.setting("switchbarPosition", estString) = "Top"
40        Else
50            settings.setting("switchbarPosition", estString) = "Bottom"
60        End If
End Property

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
End Property

Friend Function addTab(parent As IWindow, window As IWindow, order As eSwitchbarOrder, _
    caption As String, icon As CImage) As CTab
          
          Dim parentIndex As Long
          Dim tabOrder As eSwitchbarOrder

10        If Not parent Is Nothing Then
20            parentIndex = findWindowIndex(parent)
30        Else
40            parentIndex = -1
50        End If
          
          Dim insertPos As Long
          Dim count As Long
          
60        insertPos = -1
          
70        If parentIndex <> -1 Then
80            For count = parentIndex + 1 To m_tabs.count
90                If m_tabs.item(count).order > order Or m_tabs.item(count).order = eSwitchbarOrder.sboStatus Then
100                   insertPos = count
110                   Exit For
120               End If
130           Next count
140       End If
          
          Dim newTab As New CTab

150       newTab.window = window
160       newTab.caption = caption
170       newTab.icon = icon
180       newTab.order = order
          
190       m_tabs.Add newTab, insertPos
200       UserControl_Resize
          
210       Set addTab = newTab
End Function

Public Sub removeTab(aTab As CTab)
          Dim count As Long
          
10        For count = 1 To m_tabs.count
20            If m_tabs.item(count) Is aTab Then
30                m_tabs.Remove count
                  
40                If aTab Is m_selectedTab Then
50                    If count > 1 Then
60                        selectTab m_tabs.item(count - 1), True
70                    ElseIf m_tabs.count > 1 Then
80                        selectTab m_tabs.item(1), True
90                    Else
100                       Set m_selectedTab = Nothing
110                   End If
120               End If
                  
130               If aTab Is m_overTab Then
140                   Set m_overTab = Nothing
150               End If
                  
160               UserControl_Resize
                  
170               Exit Sub
180           End If
190       Next count
End Sub

Public Function findWindowIndex(window As IWindow) As Long
          Dim count As Long
          
10        For count = 1 To m_tabs.count
20            If m_tabs.item(count).window Is window Then
30                findWindowIndex = count
40                Exit Function
50            End If
60        Next count
End Function

Public Function getTabWindow(index As Long) As IWindow
10        If index > 0 And index <= m_tabs.count Then
20            Set getTabWindow = m_tabs.item(index).window
30        End If
End Function

Public Function tabCount() As Long
10        tabCount = m_tabs.count
End Function

Private Sub calculateTabWidth()
10        If m_tabs.count < 1 Then
20            Exit Sub
30        End If

          Dim controlWidth As Long
          Dim controlHeight As Long
          
40        controlWidth = UserControl.ScaleWidth - (MARGIN_LEFT + MARGIN_RIGHT)
50        controlHeight = UserControl.ScaleHeight - (MARGIN_TOP + MARGIN_BOTTOM)
          
          Dim tabWidth As Integer
          Dim tabsPerRow As Integer
          
60        tabWidth = Fix(((controlWidth * m_rows) / m_tabs.count)) - TAB_SPACING_X
          
70        If tabWidth <= m_tabHeight Then
80            m_tabWidth = m_tabHeight
90            m_tabsPerRow = Fix(controlWidth / (m_tabWidth + TAB_SPACING_X))
100           Exit Sub
110       ElseIf tabWidth >= MAX_TAB_WIDTH Then
120           tabWidth = MAX_TAB_WIDTH
130       End If
          
140       tabsPerRow = Fix(controlWidth / (tabWidth + TAB_SPACING_X))
          
          Dim removeWidth As Integer
          
150       Do While (tabsPerRow * m_rows) < m_tabs.count And tabWidth > m_tabHeight
160           removeWidth = Fix(tabWidth / m_tabs.count)
              
170           If removeWidth < 1 Then
180               tabWidth = tabWidth - 1
190           Else
200               tabWidth = tabWidth - removeWidth
210           End If
              
220           tabsPerRow = Fix(controlWidth / (tabWidth + TAB_SPACING_X))
230       Loop
          
240       If tabWidth < m_tabHeight Then
250           tabWidth = m_tabHeight
260       End If
          
270       m_tabWidth = tabWidth
280       m_tabsPerRow = Fix(controlWidth / (tabWidth + TAB_SPACING_X))
End Sub

Private Sub calcTabRects()
          Dim row As Integer
          Dim count As Integer
          Dim rowTabs As Integer
          
          Dim left As Integer
          Dim top As Integer
          
          Dim aTab As CTab
          
10        left = MARGIN_LEFT
20        top = MARGIN_TOP
          
30        For count = 1 To m_tabs.count
40            Set aTab = m_tabs.item(count)
          
50            aTab.left = left
60            aTab.right = left + m_tabWidth
70            aTab.top = top
80            aTab.bottom = top + m_tabHeight
90            aTab.updateRegion
100           aTab.visible = True
              
110           left = left + m_tabWidth + TAB_SPACING_X
          
120           rowTabs = rowTabs + 1
              
130           If rowTabs >= m_tabsPerRow Then
140               row = row + 1
                  
150               If row >= m_rows Then
                      Dim count2 As Integer
                      
160                   For count2 = count + 1 To m_tabs.count
170                       m_tabs.item(count2).visible = False
180                   Next count2
                  
190                   Exit For
200               End If
                  
210               left = MARGIN_LEFT
220               top = top + (m_tabHeight + TAB_SPACING_Y)
230               rowTabs = 0
240           End If
250       Next count
End Sub

Public Sub redrawTab(aTab As CTab)
10        aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
End Sub

Private Sub reDraw()
10        calcTabRects

          Dim backBuffer As Long
          Dim backBitmap As Long
          Dim oldBitmap As Long
          Dim oldFont As Long
          
20        backBuffer = CreateCompatibleDC(UserControl.hdc)
30        backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, _
              UserControl.ScaleHeight)
          
40        oldBitmap = SelectObject(backBuffer, backBitmap)
50        oldFont = SelectObject(backBuffer, GetCurrentObject(UserControl.hdc, OBJ_FONT))
          
          Dim controlRect As RECT
          
60        controlRect.right = UserControl.ScaleWidth
70        controlRect.bottom = UserControl.ScaleHeight
          
          Dim fillBrush As Long
          
80        If g_initialized Then
90            fillBrush = colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
100       End If
          
110       FillRect backBuffer, controlRect, fillBrush
          
          Dim count As Integer
          
          Dim x As Long
          Dim y As Long
          
120       For count = 1 To m_tabs.count
130           If m_tabs.item(count).visible Then
140               m_tabs.item(count).render backBuffer, m_tabs.item(count).left, m_tabs.item(count).top, _
                      m_tabWidth, m_tabHeight
150           End If
160       Next count
          
170       BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, _
              vbSrcCopy
              
180       SelectObject backBuffer, oldBitmap
190       SelectObject backBuffer, oldFont
              
200       DeleteObject backBitmap
210       DeleteDC backBuffer
End Sub

Private Sub mnuTabSwitch_Click()
10        selectTab m_contextTab, True
End Sub

Private Sub mnuTabClose_Click()
10        RaiseEvent closeRequest(m_contextTab)
End Sub

Private Sub UserControl_Initialize()
          Dim textMetrics As TEXTMETRIC
          
10        GetTextMetrics UserControl.hdc, textMetrics
          
20        m_tabHeight = textMetrics.tmHeight + TAB_MARGIN_TOP + TAB_MARGIN_BOTTOM

          'm_position = sbpTop
          'updateRowCount 1
          
30        m_timerFlash = 1
40        SetTimer UserControl.hwnd, m_timerFlash, 500, 0
          
50        AttachMessage Me, UserControl.hwnd, WM_TIMER
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If Not m_resizing And Not m_moving Then
20            If x <= MARGIN_LEFT And x >= 0 And y >= MARGIN_TOP And y <= (m_tabHeight + MARGIN_TOP) Then
              
30                UserControl.MousePointer = vbSizeAll
40                Exit Sub
50            End If
          
60            Select Case m_position
                  Case sbpTop
70                    If y >= UserControl.ScaleHeight - MARGIN_BOTTOM And y <= UserControl.ScaleHeight _
                          Then
                          
80                        UserControl.MousePointer = vbSizeNS
90                        Exit Sub
100                   End If
110               Case sbpBottom
120                   If y <= MARGIN_TOP And y >= 0 Then
                          
130                       UserControl.MousePointer = vbSizeNS
140                       Exit Sub
150                   End If
160           End Select
170       End If
          
180       If m_resizing Then
190           If m_position = sbpTop Then
200               processResize x, y + 5
210           ElseIf m_position = sbpBottom Then
220               processResize x, y - 5
230           End If
              
240           Exit Sub
250       End If
          
260       If m_moving Then
270           RaiseEvent moveRequest(x, y)
280           Exit Sub
290       End If
          
300       UserControl.MousePointer = vbNormal
          
310       If x < MARGIN_LEFT Or x > UserControl.ScaleWidth - MARGIN_RIGHT Or y < MARGIN_TOP Or y > _
              UserControl.ScaleHeight - MARGIN_BOTTOM Then
                  
320           If Not m_overTab Is Nothing Then
330               clearMouseOver
340           End If
              
350           ReleaseCapture
360           Exit Sub
370       End If

380       SetCapture UserControl.hwnd
390       calcMouseOver x, y
End Sub

Private Sub calcMouseOver(x As Single, y As Single)
          Dim row As Integer
          Dim tabIndex As Integer

10        row = Fix((y - MARGIN_TOP) / (m_tabHeight + TAB_SPACING_Y)) + 1
20        tabIndex = m_tabsPerRow * (row - 1)
30        tabIndex = tabIndex + Fix((x - MARGIN_LEFT) / (m_tabWidth + TAB_SPACING_X)) + 1
          
40        If tabIndex > 0 And tabIndex <= m_tabs.count Then
              Dim aTab As CTab
              
50            Set aTab = m_tabs.item(tabIndex)
                  
60            If aTab.visible And aTab.mouseOverTab(CLng(x), CLng(y)) Then
70                If Not m_overTab Is aTab Then
80                    If Not m_overTab Is Nothing Then
90                        clearMouseOver
100                   End If
                  
110                   Set m_overTab = aTab
                      
120                   If Not aTab.state = stsSelected Then
130                       aTab.state = stsMouseOver
140                       aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
150                   End If
160               End If
170           Else
180               clearMouseOver
190           End If
200       Else
210           clearMouseOver
220       End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If x <= MARGIN_LEFT And x >= 0 And y >= MARGIN_TOP And y <= (m_tabHeight + MARGIN_TOP) Then
          
20            m_moving = True
30            SetCapture UserControl.hwnd
40            Exit Sub
50        End If

60        Select Case m_position
              Case sbpTop
70                If y >= UserControl.ScaleHeight - MARGIN_BOTTOM And y <= UserControl.ScaleHeight Then
                      
80                    m_resizing = True
90                    SetCapture UserControl.hwnd
100                   Exit Sub
110               End If
120           Case sbpBottom
130               If y <= MARGIN_TOP And y >= 0 Then
                      
140                   m_resizing = True
150                   SetCapture UserControl.hwnd
160                   Exit Sub
170               End If
180       End Select

190       If Not m_overTab Is Nothing Then
200           If Button = vbKeyLButton Then
210               calcMouseOver x, y
                  
220               If Not m_overTab Is Nothing Then
230                   selectTab m_overTab, True
240               End If
250           ElseIf Button = vbKeyRButton Then
260               calcMouseOver x, y
                  
270               If Not m_overTab Is Nothing Then
280                   Set m_contextTab = m_overTab
290                   PopupMenu mnuTab
300                   Set m_contextTab = Nothing
310               End If
320           End If
330       End If
End Sub

Public Sub selectTab(aTab As CTab, events As Boolean)
10        If m_selectedTab Is aTab Then
20            Exit Sub
30        End If

40        If Not m_selectedTab Is Nothing Then
50            m_selectedTab.state = stsNormal
60            m_selectedTab.render UserControl.hdc, m_selectedTab.left, m_selectedTab.top, m_tabWidth, _
                  m_tabHeight
70        End If
          
80        Set m_selectedTab = aTab
90        aTab.state = stsSelected
          
100       aTab.activityState = tasNormal
110       aTab.trans = False
          
120       aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
          
130       If events Then
140           RaiseEvent tabSelected(aTab)
150       End If
End Sub

Public Sub tabActivity(aTab As CTab, state As eTabActivityState)
10        If aTab.state = stsNormal Or aTab.state = stsMouseOver Then
20            If state > aTab.activityState Then
30                aTab.activityState = state
40                aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
50            End If
60        End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If m_resizing Then
20            m_resizing = False
30            ReleaseCapture
40            UserControl.MousePointer = vbNormal
50        End If
          
60        If m_moving Then
70            m_moving = False
80            ReleaseCapture
90            UserControl.MousePointer = vbNormal
100       End If
End Sub

Private Sub processResize(x As Single, y As Single)
          Dim rows As Long

10        Select Case m_position
              Case sbpTop
20                rows = Fix((y - (MARGIN_TOP + MARGIN_BOTTOM)) / (m_tabHeight + TAB_SPACING_Y))
                  
30                If rows <> m_rows Then
40                    updateRowCount rows
50                End If
60            Case sbpBottom
70                rows = Fix(((UserControl.ScaleHeight + -y) - (MARGIN_TOP + MARGIN_BOTTOM)) / _
                      (m_tabHeight + TAB_SPACING_Y))
                  
80                If rows <> m_rows Then
90                    updateRowCount rows
100               End If
110       End Select
End Sub

Public Function getRequiredHeight() As Long
10        getRequiredHeight = (m_rows * (m_tabHeight + TAB_SPACING_Y)) - TAB_SPACING_Y + MARGIN_TOP + _
              MARGIN_BOTTOM
End Function


Public Property Get getMaxRows(height As Long)
10        getMaxRows = Fix((height - (MARGIN_TOP + MARGIN_BOTTOM)) / (m_tabHeight + TAB_SPACING_Y))
End Property

Private Sub updateRowCount(rows As Long)
10        If rows < 1 Then
20            m_rows = 1
30        Else
40            m_rows = rows
50        End If
          
          Dim newHeight As Long
          
60        newHeight = (m_rows * (m_tabHeight + TAB_SPACING_Y)) - TAB_SPACING_Y + MARGIN_TOP + _
              MARGIN_BOTTOM
          
70        settings.setting("switchbarRows", estNumber) = m_rows
              
80        RaiseEvent changeHeight(newHeight)
End Sub

Private Sub clearMouseOver()
10        If Not m_overTab Is Nothing Then
20            If Not m_overTab Is m_selectedTab Then
30                m_overTab.state = stsNormal
40                m_overTab.render UserControl.hdc, m_overTab.left, m_overTab.top, m_tabWidth, m_tabHeight
50            End If
                      
60            Set m_overTab = Nothing
70        End If
End Sub

Private Sub UserControl_Paint()
10        reDraw
End Sub

Private Sub UserControl_Resize()
10        calculateTabWidth
20        reDraw
End Sub

Private Sub UserControl_Terminate()
10        DetachMessage Me, UserControl.hwnd, WM_TIMER
20        KillTimer UserControl.hwnd, m_timerFlash
End Sub
