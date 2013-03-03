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

Private m_fontmanager As CFontManager

Public Event changeHeight(newHeight As Long)
Public Event moveRequest(x As Single, y As Single)
Public Event tabSelected(selectedTab As CTab)
Public Event closeRequest(aTab As CTab)

Public Property Get rows() As Long
    rows = m_rows
End Property

Public Property Let rows(newValue As Long)
    updateRowCount newValue
End Property

Private Sub IColourUser_coloursUpdated()
    reDraw
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
        Case WM_TIMER
            If wParam = m_timerFlash Then
                Dim count As Long
                
                m_trans = Not m_trans
                
                For count = 1 To m_tabs.count
                    If m_tabs.item(count).activityState = eTabActivityState.tasEvent Then
                        If settings.setting("switchbarFlashEvent", estBoolean) Then
                            updateTabFlash m_tabs.item(count), m_trans
                        End If
                    ElseIf m_tabs.item(count).activityState = eTabActivityState.tasMessage Then
                        If settings.setting("switchbarFlashMessage", estBoolean) Then
                            updateTabFlash m_tabs.item(count), m_trans
                        End If
                    ElseIf m_tabs.item(count).activityState = eTabActivityState.tasAlert Then
                        If settings.setting("switchbarFlashAlert", estBoolean) Then
                            updateTabFlash m_tabs.item(count), m_trans
                        End If
                    ElseIf m_tabs.item(count).activityState = eTabActivityState.tasHighlight Then
                        If settings.setting("switchbarFlashHighlight", estBoolean) Then
                            updateTabFlash m_tabs.item(count), m_trans
                        End If
                    End If
                Next count
            End If
    End Select
End Function

Private Sub updateTabFlash(tabItem As CTab, trans As Boolean)
    tabItem.trans = trans
    redrawTab tabItem
End Sub

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Public Property Get position() As eSwitchbarPosition
    position = m_position
End Property

Public Property Let position(newValue As eSwitchbarPosition)
    m_position = newValue
    
    If m_position = eSwitchbarPosition.sbpTop Then
        settings.setting("switchbarPosition", estString) = "Top"
    Else
        settings.setting("switchbarPosition", estString) = "Bottom"
    End If
End Property

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontmanager = RHS
End Property

Friend Function addTab(parent As IWindow, window As IWindow, order As eSwitchbarOrder, caption As String, icon As CImage) As CTab
    
    Dim parentIndex As Long
    Dim tabOrder As eSwitchbarOrder

    If Not parent Is Nothing Then
        parentIndex = findWindowIndex(parent)
    Else
        parentIndex = -1
    End If
    
    Dim insertPos As Long
    Dim count As Long
    
    insertPos = -1
    
    If parentIndex <> -1 Then
        For count = parentIndex + 1 To m_tabs.count
            If m_tabs.item(count).order > order Or m_tabs.item(count).order = eSwitchbarOrder.sboStatus Then
                insertPos = count
                Exit For
            End If
        Next count
    End If
    
    Dim newTab As New CTab

    newTab.window = window
    newTab.caption = caption
    newTab.icon = icon
    newTab.order = order
    
    m_tabs.Add newTab, insertPos
    UserControl_Resize
    
    Set addTab = newTab
End Function

Public Sub removeTab(aTab As CTab)
    Dim count As Long
    
    For count = 1 To m_tabs.count
        If m_tabs.item(count) Is aTab Then
            m_tabs.Remove count
            
            If aTab Is m_selectedTab Then
                If count > 1 Then
                    selectTab m_tabs.item(count - 1), True
                ElseIf m_tabs.count > 1 Then
                    selectTab m_tabs.item(1), True
                Else
                    Set m_selectedTab = Nothing
                End If
            End If
            
            If aTab Is m_overTab Then
                Set m_overTab = Nothing
            End If
            
            UserControl_Resize
            
            Exit Sub
        End If
    Next count
End Sub

Public Function findWindowIndex(window As IWindow) As Long
    Dim count As Long
    
    For count = 1 To m_tabs.count
        If m_tabs.item(count).window Is window Then
            findWindowIndex = count
            Exit Function
        End If
    Next count
End Function

Public Function getTabWindow(index As Long) As IWindow
    If index > 0 And index <= m_tabs.count Then
        Set getTabWindow = m_tabs.item(index).window
    End If
End Function

Public Function tabCount() As Long
    tabCount = m_tabs.count
End Function

Private Sub calculateTabWidth()
    If m_tabs.count < 1 Then
        Exit Sub
    End If

    Dim controlWidth As Long
    Dim controlHeight As Long
    
    controlWidth = UserControl.ScaleWidth - (MARGIN_LEFT + MARGIN_RIGHT)
    controlHeight = UserControl.ScaleHeight - (MARGIN_TOP + MARGIN_BOTTOM)
    
    Dim tabWidth As Integer
    Dim tabsPerRow As Integer
    
    tabWidth = Fix(((controlWidth * m_rows) / m_tabs.count)) - TAB_SPACING_X
    
    If tabWidth <= m_tabHeight Then
        m_tabWidth = m_tabHeight
        m_tabsPerRow = Fix(controlWidth / (m_tabWidth + TAB_SPACING_X))
        Exit Sub
    ElseIf tabWidth >= MAX_TAB_WIDTH Then
        tabWidth = MAX_TAB_WIDTH
    End If
    
    tabsPerRow = Fix(controlWidth / (tabWidth + TAB_SPACING_X))
    
    Dim removeWidth As Integer
    
    Do While (tabsPerRow * m_rows) < m_tabs.count And tabWidth > m_tabHeight
        removeWidth = Fix(tabWidth / m_tabs.count)
        
        If removeWidth < 1 Then
            tabWidth = tabWidth - 1
        Else
            tabWidth = tabWidth - removeWidth
        End If
        
        tabsPerRow = Fix(controlWidth / (tabWidth + TAB_SPACING_X))
    Loop
    
    If tabWidth < m_tabHeight Then
        tabWidth = m_tabHeight
    End If
    
    m_tabWidth = tabWidth
    m_tabsPerRow = Fix(controlWidth / (tabWidth + TAB_SPACING_X))
End Sub

Private Sub calcTabRects()
    Dim row As Integer
    Dim count As Integer
    Dim rowTabs As Integer
    
    Dim left As Integer
    Dim top As Integer
    
    Dim aTab As CTab
    
    left = MARGIN_LEFT
    top = MARGIN_TOP
    
    For count = 1 To m_tabs.count
        Set aTab = m_tabs.item(count)
    
        aTab.left = left
        aTab.right = left + m_tabWidth
        aTab.top = top
        aTab.bottom = top + m_tabHeight
        aTab.updateRegion
        aTab.visible = True
        
        left = left + m_tabWidth + TAB_SPACING_X
    
        rowTabs = rowTabs + 1
        
        If rowTabs >= m_tabsPerRow Then
            row = row + 1
            
            If row >= m_rows Then
                Dim count2 As Integer
                
                For count2 = count + 1 To m_tabs.count
                    m_tabs.item(count2).visible = False
                Next count2
            
                Exit For
            End If
            
            left = MARGIN_LEFT
            top = top + (m_tabHeight + TAB_SPACING_Y)
            rowTabs = 0
        End If
    Next count
End Sub

Public Sub redrawTab(aTab As CTab)
    aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
End Sub

Private Sub reDraw()
    calcTabRects

    Dim backBuffer As Long
    Dim backBitmap As Long
    Dim oldBitmap As Long
    Dim oldFont As Long
    
    backBuffer = CreateCompatibleDC(UserControl.hdc)
    backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
    
    oldBitmap = SelectObject(backBuffer, backBitmap)
    oldFont = SelectObject(backBuffer, GetCurrentObject(UserControl.hdc, OBJ_FONT))
    
    Dim controlRect As RECT
    
    controlRect.right = UserControl.ScaleWidth
    controlRect.bottom = UserControl.ScaleHeight
    
    Dim fillBrush As Long
    
    If g_initialized Then
        fillBrush = colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    End If
    
    FillRect backBuffer, controlRect, fillBrush
    
    Dim count As Integer
    
    Dim x As Long
    Dim y As Long
    
    For count = 1 To m_tabs.count
        If m_tabs.item(count).visible Then
            m_tabs.item(count).render backBuffer, m_tabs.item(count).left, m_tabs.item(count).top, m_tabWidth, m_tabHeight
        End If
    Next count
    
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, vbSrcCopy
        
    SelectObject backBuffer, oldBitmap
    SelectObject backBuffer, oldFont
        
    DeleteObject backBitmap
    DeleteDC backBuffer
End Sub

Private Sub mnuTabSwitch_Click()
    selectTab m_contextTab, True
End Sub

Private Sub mnuTabClose_Click()
    RaiseEvent closeRequest(m_contextTab)
End Sub

Private Sub UserControl_Initialize()
    Dim textMetrics As TEXTMETRIC
    
    GetTextMetrics UserControl.hdc, textMetrics
    
    m_tabHeight = textMetrics.tmHeight + TAB_MARGIN_TOP + TAB_MARGIN_BOTTOM

    'm_position = sbpTop
    'updateRowCount 1
    
    m_timerFlash = 1
    SetTimer UserControl.hwnd, m_timerFlash, 500, 0
    
    AttachMessage Me, UserControl.hwnd, WM_TIMER
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not m_resizing And Not m_moving Then
        If x <= MARGIN_LEFT And x >= 0 And y >= MARGIN_TOP And y <= (m_tabHeight + MARGIN_TOP) Then
        
            UserControl.MousePointer = vbSizeAll
            Exit Sub
        End If
    
        Select Case m_position
            Case sbpTop
                If y >= UserControl.ScaleHeight - MARGIN_BOTTOM And y <= UserControl.ScaleHeight Then
                    
                    UserControl.MousePointer = vbSizeNS
                    Exit Sub
                End If
            Case sbpBottom
                If y <= MARGIN_TOP And y >= 0 Then
                    
                    UserControl.MousePointer = vbSizeNS
                    Exit Sub
                End If
        End Select
    End If
    
    If m_resizing Then
        If m_position = sbpTop Then
            processResize x, y + 5
        ElseIf m_position = sbpBottom Then
            processResize x, y - 5
        End If
        
        Exit Sub
    End If
    
    If m_moving Then
        RaiseEvent moveRequest(x, y)
        Exit Sub
    End If
    
    UserControl.MousePointer = vbNormal
    
    If x < MARGIN_LEFT Or x > UserControl.ScaleWidth - MARGIN_RIGHT Or y < MARGIN_TOP Or y > UserControl.ScaleHeight - MARGIN_BOTTOM Then
            
        If Not m_overTab Is Nothing Then
            clearMouseOver
        End If
        
        ReleaseCapture
        Exit Sub
    End If

    SetCapture UserControl.hwnd
    calcMouseOver x, y
End Sub

Private Sub calcMouseOver(x As Single, y As Single)
    Dim row As Integer
    Dim tabIndex As Integer

    row = Fix((y - MARGIN_TOP) / (m_tabHeight + TAB_SPACING_Y)) + 1
    tabIndex = m_tabsPerRow * (row - 1)
    tabIndex = tabIndex + Fix((x - MARGIN_LEFT) / (m_tabWidth + TAB_SPACING_X)) + 1
    
    If tabIndex > 0 And tabIndex <= m_tabs.count Then
        Dim aTab As CTab
        
        Set aTab = m_tabs.item(tabIndex)
            
        If aTab.visible And aTab.mouseOverTab(CLng(x), CLng(y)) Then
            If Not m_overTab Is aTab Then
                If Not m_overTab Is Nothing Then
                    clearMouseOver
                End If
            
                Set m_overTab = aTab
                
                If Not aTab.state = stsSelected Then
                    aTab.state = stsMouseOver
                    aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
                End If
            End If
        Else
            clearMouseOver
        End If
    Else
        clearMouseOver
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x <= MARGIN_LEFT And x >= 0 And y >= MARGIN_TOP And y <= (m_tabHeight + MARGIN_TOP) Then
    
        m_moving = True
        SetCapture UserControl.hwnd
        Exit Sub
    End If

    Select Case m_position
        Case sbpTop
            If y >= UserControl.ScaleHeight - MARGIN_BOTTOM And y <= UserControl.ScaleHeight Then
                
                m_resizing = True
                SetCapture UserControl.hwnd
                Exit Sub
            End If
        Case sbpBottom
            If y <= MARGIN_TOP And y >= 0 Then
                
                m_resizing = True
                SetCapture UserControl.hwnd
                Exit Sub
            End If
    End Select

    If Not m_overTab Is Nothing Then
        If Button = vbKeyLButton Then
            calcMouseOver x, y
            
            If Not m_overTab Is Nothing Then
                selectTab m_overTab, True
            End If
        ElseIf Button = vbKeyRButton Then
            calcMouseOver x, y
            
            If Not m_overTab Is Nothing Then
                Set m_contextTab = m_overTab
                PopupMenu mnuTab
                Set m_contextTab = Nothing
            End If
        End If
    End If
End Sub

Public Sub selectTab(aTab As CTab, events As Boolean)
    If m_selectedTab Is aTab Then
        Exit Sub
    End If

    If Not m_selectedTab Is Nothing Then
        m_selectedTab.state = stsNormal
        m_selectedTab.render UserControl.hdc, m_selectedTab.left, m_selectedTab.top, m_tabWidth, m_tabHeight
    End If
    
    Set m_selectedTab = aTab
    aTab.state = stsSelected
    
    aTab.activityState = tasNormal
    aTab.trans = False
    
    aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
    
    If events Then
        RaiseEvent tabSelected(aTab)
    End If
End Sub

Public Sub tabActivity(aTab As CTab, state As eTabActivityState)
    If aTab.state = stsNormal Or aTab.state = stsMouseOver Then
        If state > aTab.activityState Then
            aTab.activityState = state
            aTab.render UserControl.hdc, aTab.left, aTab.top, m_tabWidth, m_tabHeight
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_resizing Then
        m_resizing = False
        ReleaseCapture
        UserControl.MousePointer = vbNormal
    End If
    
    If m_moving Then
        m_moving = False
        ReleaseCapture
        UserControl.MousePointer = vbNormal
    End If
End Sub

Private Sub processResize(x As Single, y As Single)
    Dim rows As Long

    Select Case m_position
        Case sbpTop
            rows = Fix((y - (MARGIN_TOP + MARGIN_BOTTOM)) / (m_tabHeight + TAB_SPACING_Y))
            
            If rows <> m_rows Then
                updateRowCount rows
            End If
        Case sbpBottom
            rows = Fix(((UserControl.ScaleHeight + -y) - (MARGIN_TOP + MARGIN_BOTTOM)) / (m_tabHeight + TAB_SPACING_Y))
            
            If rows <> m_rows Then
                updateRowCount rows
            End If
    End Select
End Sub

Public Function getRequiredHeight() As Long
    getRequiredHeight = (m_rows * (m_tabHeight + TAB_SPACING_Y)) - TAB_SPACING_Y + MARGIN_TOP + MARGIN_BOTTOM
End Function


Public Property Get getMaxRows(height As Long)
    getMaxRows = Fix((height - (MARGIN_TOP + MARGIN_BOTTOM)) / (m_tabHeight + TAB_SPACING_Y))
End Property

Private Sub updateRowCount(rows As Long)
    If rows < 1 Then
        m_rows = 1
    Else
        m_rows = rows
    End If
    
    Dim newHeight As Long
    
    newHeight = (m_rows * (m_tabHeight + TAB_SPACING_Y)) - TAB_SPACING_Y + MARGIN_TOP + MARGIN_BOTTOM
    
    settings.setting("switchbarRows", estNumber) = m_rows
        
    RaiseEvent changeHeight(newHeight)
End Sub

Private Sub clearMouseOver()
    If Not m_overTab Is Nothing Then
        If Not m_overTab Is m_selectedTab Then
            m_overTab.state = stsNormal
            m_overTab.render UserControl.hdc, m_overTab.left, m_overTab.top, m_tabWidth, m_tabHeight
        End If
                
        Set m_overTab = Nothing
    End If
End Sub

Private Sub UserControl_Paint()
    reDraw
End Sub

Private Sub UserControl_Resize()
    calculateTabWidth
    reDraw
End Sub

Private Sub UserControl_Terminate()
    DetachMessage Me, UserControl.hwnd, WM_TIMER
    KillTimer UserControl.hwnd, m_timerFlash
End Sub
