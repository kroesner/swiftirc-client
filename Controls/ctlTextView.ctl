VERSION 5.00
Begin VB.UserControl ctlTextView 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   99  'Custom
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   708
End
Attribute VB_Name = "ctlTextView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eTvEventFlags
    TVE_NONE = &H0
    TVE_NOEVENT = &H1 'Only has meaning when applied to individual events/lines
    TVE_VISIBLE = &H2
    TVE_TIMESTAMP = &H4
    TVE_INDENTWRAP = &H8
    TVE_CUSTOMIRCCOLOUR = &H10 'Only affects raw text
    TVE_SEPERATE_TOP = &H20
    TVE_SEPERATE_BOTTOM = &H40
    TVE_SEPERATE_BOTH = TVE_SEPERATE_TOP Or TVE_SEPERATE_BOTTOM
    TVE_SEPERATE_EXPLICIT = &H80
    TVE_STANDARD = TVE_VISIBLE Or TVE_TIMESTAMP
    TVE_USERTEXT = TVE_VISIBLE Or TVE_TIMESTAMP Or TVE_INDENTWRAP
End Enum

Implements ISubclass
Implements IWindow
Implements IFontUser
Implements IColourUser

Private m_atBottom As Boolean

Private m_mouseX As Single
Private m_mouseY As Single

Private m_enableLogging As Boolean
Private m_logBaseName As String
Private m_logFilename As String
Private m_logHandle As Long

Private Const MAX_LINES As Long = 5000
Private Const MAX_REMOVE_LINES As Long = 5
Private Const BOTTOM_BUFFER As Long = 5

Private m_realWindow As VBControlExtender
Private m_selecting As Boolean

Private m_realSelectStartX As Integer
Private m_realSelectStartY As Integer

Private m_oldSelectStartY As Integer
Private m_oldSelectEndY As Integer

Private m_selectStartX As Integer
Private m_selectStartY As Integer
Private m_selectEndX As Integer
Private m_selectEndY As Integer
Private m_moveX As Integer

Private m_backBuffer As Long
Private m_backBitmap As Long
Private m_drawingData As New CDrawingData

Private m_eventManager As CEventManager

Private m_allLines As New cArrayList
Private m_lines As New cArrayList
Private m_visibleLines As New cArrayList

Private m_fontHeight As Long
Private m_pageLines As Long
Private m_hScrollBar As Long

Private m_currentVirtLine As Long
Private m_currentPhysLine As Long

Private m_topRealLine As Long
Private m_currentRealLine As Long
Private m_realLineCount As Long

Public Event doubleClick()
Public Event noLongerNeedFocus()
Public Event clickedUrl(url As String)

Private m_wasOverUrl As Boolean
Private m_url As String

Private m_scrollbarEnabled As Boolean

Private m_fontmanager As CFontManager

Private Sub scrolled()
    If m_currentRealLine >= m_realLineCount Then
        m_atBottom = True
    Else
        m_atBottom = False
    End If
    
    If m_wasOverUrl Then
        UserControl.MousePointer = 0
        m_wasOverUrl = False
    End If
End Sub

Public Property Get logName() As String
    logName = m_logBaseName
End Property

Public Property Let logName(newValue As String)
    Dim oldName As String
    
    oldName = m_logBaseName
    m_logBaseName = newValue
    
    If m_logBaseName <> oldName And m_enableLogging Then
        closeLog
    End If
End Property

Public Property Get enableLogging() As Boolean
    enableLogging = m_enableLogging
End Property

Public Property Let enableLogging(newValue As Boolean)
    m_enableLogging = newValue
    
    If Not m_enableLogging Then
        closeLog
    End If
End Property

Public Property Get ignoreSeperators() As Boolean
    ignoreSeperators = m_drawingData.ignoreSeperators
End Property

Public Property Let ignoreSeperators(newValue As Boolean)
    m_drawingData.ignoreSeperators = newValue
End Property

Public Sub clear()
    m_currentRealLine = 1
    m_currentVirtLine = 1
    m_currentPhysLine = 1
    m_realLineCount = 1

    scrolled

    m_visibleLines.clear
    m_lines.clear
    m_allLines.clear
    refresh
End Sub

Private Sub updateColours()
    m_drawingData.defaultForeColour = g_textViewFore
    m_drawingData.defaultBackColour = g_textViewBack
    reDraw
End Sub

Public Property Let eventManager(newValue As CEventManager)
    Set m_eventManager = newValue
End Property

Public Property Get foreColour() As Byte
    foreColour = m_drawingData.defaultForeColour
End Property

Public Property Let foreColour(newValue As Byte)
    m_drawingData.defaultForeColour = newValue
    reDraw
End Property

Public Property Get backColour() As Byte
    backColour = m_drawingData.defaultBackColour
End Property

Public Property Let backColour(newValue As Byte)
    m_drawingData.defaultBackColour = newValue
    reDraw
End Property

Public Sub addEvent(eventName As String, params() As String)
    addEventEx eventName, Nothing, vbNullString, TVE_NONE, params
End Sub

Public Sub addEventEx(eventName As String, userStyle As CUserStyle, username As String, flags As Long, params() As String)
    Dim line As New CLine
    Dim aEvent As CEvent
    
   On Error GoTo addEventEx_Error

    Set aEvent = m_eventManager.findEvent(eventName)
    
    If aEvent Is Nothing Then
        Exit Sub
    End If
    
    line.init aEvent, flags, userStyle, username, params
    
    m_allLines.Add line
    
    If Not line.shouldShow Then
        Exit Sub
    End If
    
    If m_lines.count > 0 Then
        If m_lines.item(m_lines.count).seperatorBottom(m_drawingData.ignoreSeperators) Then
            line.seperatorAbove = True
        End If
    End If
    
    m_lines.Add line
    
    If m_atBottom Then
        m_realLineCount = m_realLineCount + 1
        
        If UserControl.Extender.visible = True Then
            wrap line, True
            scrollDown m_realLineCount - m_currentRealLine
        Else
            If m_lines.count > 1 Then
                m_currentRealLine = m_currentRealLine + 1
                m_currentVirtLine = m_lines.count
                m_currentPhysLine = 1
            End If
        End If
    Else
        m_realLineCount = m_realLineCount + 1
    End If
    
    line.wasDisplayed = True
    removeOldLines
    
    If m_realLineCount = 1 Then
        reDraw
    End If
    
    updateScrollBar
    
    logWrite line
    
   On Error GoTo 0
   Exit Sub

addEventEx_Error:
    handleError "addEventEx", Err.Number, Err.Description, Erl, eventName
End Sub

Public Sub addRawText(format As String, params() As String)
    addRawTextEx eventColours.otherText, 0, format, Nothing, vbNullString, TVE_NONE, params
End Sub

Public Sub addRawTextEx(eventColour As CEventColour, foreColour As Byte, format As String, userStyle As CUserStyle, username As String, flags As Long, params() As String)
    
    Dim line As New CLine
    
    On Error GoTo addRawTextEx_Error

    line.initEx eventColour, foreColour, format, userStyle, username, flags Or TVE_NOEVENT Or TVE_VISIBLE, params
    
    m_allLines.Add line
    m_lines.Add line
    
    If m_atBottom Then
        If UserControl.Extender.visible Then
            m_realLineCount = m_realLineCount + 1
            wrap line, True
            scrollDown m_realLineCount - m_currentRealLine
        Else
            If m_lines.count > 1 Then
                m_currentRealLine = m_currentRealLine + 1
                m_currentVirtLine = m_lines.count
                m_currentPhysLine = 1
                m_realLineCount = m_realLineCount + 1
            End If
        End If
    Else
        m_realLineCount = m_realLineCount + 1
    End If
    
    line.wasDisplayed = True

    If m_realLineCount = 1 Then
        reDraw
    End If

    removeOldLines
    updateScrollBar
    
    logWrite line

   On Error GoTo 0
   Exit Sub

addRawTextEx_Error:
    handleError "addRawTextEx", Err.Number, Err.Description, Erl, format
End Sub

Private Sub removeOldLines()
    If m_lines.count <= MAX_LINES Then
        Exit Sub
    End If

    Dim lineCount As Long
    
    lineCount = m_lines.count - MAX_LINES
    
    If lineCount > MAX_REMOVE_LINES Then
        lineCount = MAX_REMOVE_LINES
    End If

    If m_currentVirtLine <= lineCount Or m_currentRealLine <= m_pageLines + MAX_REMOVE_LINES Then
        Exit Sub
    End If
    
    If lineCount = 0 Then
        Exit Sub
    End If
    
    Dim count As Long
    
    For count = 1 To lineCount
        m_realLineCount = m_realLineCount - m_lines.item(1).physLineCount
        m_currentRealLine = m_currentRealLine - m_lines.item(1).physLineCount
        m_currentVirtLine = m_currentVirtLine - 1
        m_lines.Remove 1
    Next count
    
    Dim removed As Long
    
    For count = 1 To m_allLines.count
        If removed = lineCount Then
            Exit For
        End If
        
        If m_allLines.item(1).wasDisplayed Then
            removed = removed + 1
        End If
        
        m_allLines.Remove 1
    Next count
End Sub

Public Sub updateVisibility()
    calculateEventVisibility
End Sub

Private Sub calculateEventVisibility()
    Dim count As Integer
    Dim line As CLine
    Dim origVirtLine As Integer
    
    origVirtLine = m_currentVirtLine
    
    For count = 1 To m_lines.count
        Set line = m_lines.item(count)
    
        If Not line.shouldShow Then
            m_realLineCount = m_realLineCount - line.physLineCount
            
            If count <= origVirtLine Then
                m_currentRealLine = m_currentRealLine - line.physLineCount
                m_currentVirtLine = m_currentVirtLine - 1
            End If
            
            If line.wrapped Then
                line.clearWrap
            End If
            
            line.wasDisplayed = False
        End If
    Next count
    
    m_lines.clear
    
    Dim linesAdded As Boolean
    
    For count = 1 To m_allLines.count
        Set line = m_allLines.item(count)
        
        If line.shouldShow Then
            If Not line.wasDisplayed Then
                linesAdded = True
            
                m_realLineCount = m_realLineCount + line.physLineCount
                
                If count <= m_currentVirtLine Then
                    m_currentRealLine = m_currentRealLine + line.physLineCount
                    m_currentVirtLine = m_currentVirtLine + 1
                End If
                
                line.wasDisplayed = True
            End If
            
            m_lines.Add line
        End If
    Next count
    
    If linesAdded Then
        If m_currentVirtLine = 0 Then
            m_currentVirtLine = 1
            m_currentPhysLine = 1
            m_currentRealLine = 1
        End If
    End If
    
    If m_lines.count > 0 Then
        If m_currentPhysLine > m_lines.item(m_currentVirtLine).physLineCount Then
            m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount
        End If
    End If
    
    refresh
End Sub

Private Sub IColourUser_coloursUpdated()
    m_drawingData.setPalette colourThemes.currentTheme.getPalette
    m_drawingData.defaultBackColour = g_textViewBack
    m_drawingData.defaultForeColour = g_textViewFore
    hardRedraw
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontmanager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
    m_fontHeight = m_fontmanager.fontHeight
    m_drawingData.fontHeight = m_fontHeight
    m_drawingData.fontManager = m_fontmanager
    refresh
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_SYSCOLORCHANGE
            ISubclass_MsgResponse = emrPreprocess
        Case Else
            ISubclass_MsgResponse = emrConsume
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
        Case WM_VSCROLL
            Dim scrollCode As Integer
            
            scrollCode = LoWord(wParam)
            processScroll scrollCode
        Case WM_MOUSEWHEEL
            Dim accumDelta As Integer
            
            accumDelta = HiWord(wParam)
            processMouseWheel accumDelta
    End Select
End Function

Public Sub processMouseWheel(ByVal accumDelta As Integer)
    Dim lines As Long
    
    Do While accumDelta >= 40
        lines = lines + 1
        accumDelta = accumDelta - 40
    Loop
    
    If lines > 0 Then
        If lines >= m_currentRealLine Then
            lines = m_currentRealLine - 1
        End If
        
        scrollUp lines
    End If
    
    lines = 0
    
    Do While accumDelta <= -40
        lines = lines + 1
        accumDelta = accumDelta + 40
    Loop
    
    If lines > 0 Then
        If lines > m_realLineCount - m_currentRealLine Then
            lines = m_realLineCount - m_currentRealLine
        End If
        
        scrollDown lines
    End If
    
    updateScrollBar
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub UserControl_DblClick()
    RaiseEvent doubleClick
End Sub

Private Sub UserControl_Initialize()
    m_logHandle = -1

    m_fontHeight = 20

    m_drawingData.setPalette colourThemes.currentTheme.getPalette
    m_drawingData.fontHeight = m_fontHeight
    m_drawingData.fontManager = m_fontmanager

    initScrollBar
    initMessages

    m_currentRealLine = 1
    m_currentVirtLine = 1
    m_currentPhysLine = 1
    
    m_atBottom = True
    
    updateScrollBar
    
    updateColours
End Sub

Private Sub initMessages()
    AttachMessage Me, UserControl.hwnd, WM_CTLCOLORSCROLLBAR
    AttachMessage Me, UserControl.hwnd, WM_VSCROLL
    AttachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
    AttachMessage Me, UserControl.hwnd, WM_SYSCOLORCHANGE
End Sub

Private Sub deInitMessages()
    DetachMessage Me, UserControl.hwnd, WM_CTLCOLORSCROLLBAR
    DetachMessage Me, UserControl.hwnd, WM_VSCROLL
    DetachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
    DetachMessage Me, UserControl.hwnd, WM_SYSCOLORCHANGE
End Sub

Private Sub scrollTop()
    m_currentVirtLine = 1
    m_currentPhysLine = 1
    m_currentRealLine = 1
    
    scrolled
    
    updateScrollBar
    reDraw
End Sub

Private Sub scrollBottom()
    Dim line As CLine
    
    m_currentVirtLine = m_lines.count
    
    Set line = m_lines.item(m_currentVirtLine)
    
    wrap line, True
    
    m_currentPhysLine = line.physLineCount
    m_currentRealLine = m_realLineCount
    
    scrolled
    
    updateScrollBar
    reDraw
End Sub

Private Sub scrollUp(lines As Long)
   On Error GoTo scrollUp_Error

    If lines < 1 Then
        Exit Sub
    End If
    
    Dim count As Long
    Dim lastVisibleLine As Long
    Dim visibleLines As Long
    Dim line As CLine
    
    For count = m_currentVirtLine To 1 Step -1
        Set line = m_lines.item(count)
   
        If count = m_currentVirtLine Then
            visibleLines = visibleLines + m_currentPhysLine
        Else
            visibleLines = visibleLines + line.physLineCount
        End If
   
        If visibleLines > m_pageLines Then
            If count < m_currentVirtLine Then
                visibleLines = visibleLines - line.physLineCount
            End If
            
            Exit For
        End If
    Next count
    
    lastVisibleLine = count
    
    For count = lines To 1 Step -1
        If m_currentPhysLine > 1 Then
            m_currentPhysLine = m_currentPhysLine - 1
        Else
            If m_currentVirtLine = 1 Then
                lines = lines - count
                Exit For
            End If
            
            If m_visibleLines.count > 0 Then
                If m_visibleLines.item(m_visibleLines.count).selected Then
                    m_visibleLines.item(m_visibleLines.count).unSelect
                End If
            
                m_visibleLines.item(m_visibleLines.count).clearWrap
                m_visibleLines.item(m_visibleLines.count).visible = False
                
                m_visibleLines.Remove m_visibleLines.count
            End If
            
            m_currentVirtLine = m_currentVirtLine - 1
            wrap m_lines.item(m_currentVirtLine), True
            m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount
        End If
    Next count
    
    If lines < 1 Then
        Exit Sub
    End If
    
    m_currentRealLine = m_currentRealLine - lines
    
    If lines >= m_pageLines Then
        reDraw
        Exit Sub
    End If

    Dim copyHeight As Long
    Dim topMargin As Long
    Dim copyY As Long
    
    topMargin = UserControl.ScaleHeight - (visibleLines * m_fontHeight)
    copyHeight = m_fontHeight * (visibleLines - lines)
    copyY = topMargin + (m_fontHeight * lines)
    
    BitBlt m_backBuffer, 0, copyY, UserControl.ScaleWidth - SB_WIDTH, copyHeight, m_backBuffer, 0, topMargin, vbSrcCopy
    
    drawLines lastVisibleLine, 0, CInt(copyY)
    displayBackBuffer
    
    scrolled

   On Error GoTo 0
   Exit Sub

scrollUp_Error:
    handleError "scrollUp", Err.Number, Err.Description, Erl, CStr(lines) & " lines"
End Sub

Private Sub scrollDown(lines As Long)
   On Error GoTo scrollDown_Error

    If lines < 1 Then
        Exit Sub
    End If
    
    Dim count As Long
    Dim diff As Long
    
    For count = lines To 1 Step -1
        If m_lines.item(m_currentVirtLine).physLineCount > m_currentPhysLine Then
            m_currentPhysLine = m_currentPhysLine + 1
        Else
            If m_lines.count <= m_currentVirtLine Then
                lines = lines - count
                Exit For
            End If
            
            m_currentVirtLine = m_currentVirtLine + 1
            m_currentPhysLine = 1
            
            wrap m_lines.item(m_currentVirtLine), True
        End If
    Next count
    
    If lines < 1 Then
        Exit Sub
    End If
    
    m_currentRealLine = m_currentRealLine + lines
    
    If lines >= m_pageLines Then
        reDraw
        Exit Sub
    End If
    
    Dim scrollDist As Long
    Dim copyHeight As Long
    
    scrollDist = m_fontHeight * lines
    copyHeight = UserControl.ScaleHeight - scrollDist
    
    BitBlt m_backBuffer, 0, 0, UserControl.ScaleWidth - SB_WIDTH, copyHeight, m_backBuffer, 0, scrollDist, vbSrcCopy
    
    For count = 1 To m_visibleLines.count
        m_visibleLines.item(count).shiftedUp scrollDist
    Next count
    
    Do While m_visibleLines.count > 0
        If m_visibleLines.item(1).bottom >= 0 Then
            Exit Do
        End If
        
        If m_visibleLines.item(1).selected Then
            m_visibleLines.item(1).unSelect
        End If
        
        m_visibleLines.item(1).visible = False
        m_visibleLines.item(1).clearWrap
        
        m_visibleLines.Remove 1
    Loop
    
    drawLines m_currentVirtLine, copyHeight, UserControl.ScaleHeight - copyHeight
    displayBackBuffer
    
    scrolled

   On Error GoTo 0
   Exit Sub

scrollDown_Error:
    handleError "scrollDown", Err.Number, Err.Description, Erl, CStr(lines) & " lines"
End Sub

Public Sub pageUp()
    If m_currentRealLine - m_pageLines < 1 Then
        scrollUp m_currentRealLine - 1
    Else
        scrollUp m_pageLines
    End If
    
    scrolled
    updateScrollBar
End Sub

Public Sub pageDown()
    If m_currentRealLine + m_pageLines > m_realLineCount Then
        scrollDown m_realLineCount - m_currentRealLine
    Else
        scrollDown m_pageLines
    End If
    
    scrolled
    updateScrollBar
End Sub

Private Sub processScroll(scrollCode As Integer)
   On Error GoTo processScroll_Error

    Select Case scrollCode
        Case SB_PAGEUP
            pageUp
        Case SB_PAGEDOWN
            pageDown
        Case SB_LINEUP
            If m_currentRealLine > 1 Then
                scrollUp 1
                updateScrollBar
            End If
        Case SB_LINEDOWN
            If m_currentRealLine < m_realLineCount Then
                scrollDown 1
                updateScrollBar
            End If
        Case SB_TOP
            scrollTop
        Case SB_BOTTOM
            scrollBottom
        Case SB_THUMBTRACK
            Dim si As SCROLLINFO
            Dim diff As Long
            
            si.cbSize = Len(si)
            si.fMask = SIF_TRACKPOS
            
            GetScrollInfo m_hScrollBar, SB_CTL, si
            
            If si.nTrackPos = 1 Then
                scrollTop
                Exit Sub
            ElseIf si.nTrackPos >= m_realLineCount Then
                scrollBottom
                Exit Sub
            End If
            
            Dim pos As Single
            Dim vline As Integer
            
            diff = m_currentRealLine - si.nTrackPos
            
            If diff < 0 Then
                If -diff > (m_pageLines) Then
                    bigScroll diff
                Else
                    If m_currentRealLine - diff > m_realLineCount Then
                        scrollDown m_realLineCount - m_currentRealLine
                        updateScrollBar
                    Else
                        scrollDown -diff
                        updateScrollBar
                    End If
                End If
            Else
                If diff > (m_pageLines) Then
                    bigScroll diff
                Else
                    If m_currentRealLine - diff < 1 Then
                        scrollUp diff - m_currentRealLine
                        updateScrollBar
                    Else
                        scrollUp diff
                        updateScrollBar
                    End If
                End If
            End If
    End Select

   On Error GoTo 0
   Exit Sub

processScroll_Error:
    handleError "processScroll", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub bigScroll(diff As Long)
    Dim count As Long

    For count = m_currentVirtLine To 1 Step -1
        If m_lines.item(count).wrapped = False Then
            Exit For
        End If
            
        m_lines.item(count).clearWrap
    Next count

    If diff > 0 Then
        For count = diff To 1 Step -1
            If m_currentPhysLine = 1 Then
                If m_currentVirtLine = 1 Then Exit For
                
                m_currentVirtLine = m_currentVirtLine - 1
                m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount
            Else
                m_currentPhysLine = m_currentPhysLine - 1
            End If
        Next count
    Else
        For count = -diff To 1 Step -1
            If m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount Then
                If m_currentVirtLine = m_lines.count Then Exit For
            
                m_currentVirtLine = m_currentVirtLine + 1
                m_currentPhysLine = 1
            Else
                m_currentPhysLine = m_currentPhysLine + 1
            End If
        Next count
    End If
    
    m_currentRealLine = m_currentRealLine - diff
    
    If m_currentRealLine < 1 Then
        m_currentRealLine = 1
    ElseIf m_currentRealLine > m_realLineCount Then
        m_currentRealLine = m_realLineCount
    End If
    
    scrolled
    
    updateScrollBar
    reDraw
End Sub

Private Sub initScrollBar()
    m_hScrollBar = CreateWindowEx(0, "SCROLLBAR", "", WS_CHILD Or SBS_VERT, UserControl.ScaleWidth - SB_WIDTH, 0, SB_WIDTH, UserControl.ScaleHeight, UserControl.hwnd, 0, App.hInstance, ByVal 0&)
    ShowScrollBar m_hScrollBar, SB_CTL, 1
    
    m_scrollbarEnabled = True
    
    updateScrollBar
End Sub

Private Sub updateScrollBar()
    Dim si As SCROLLINFO
    
    si.cbSize = Len(si)
    si.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
    
    si.nMin = 1
    
    If m_realLineCount < 2 Then
        si.nMax = 0
    Else
        si.nMax = m_realLineCount + (m_pageLines - 1)
    End If
    
    si.nPage = m_pageLines
    si.nPos = m_currentRealLine
    
    SetScrollInfo m_hScrollBar, SB_CTL, si, 0
    
    If si.nMax = 0 Then
        EnableScrollBar m_hScrollBar, SB_CTL, ESB_DISABLE_BOTH
    Else
        EnableScrollBar m_hScrollBar, SB_CTL, ESB_ENABLE_BOTH
    End If
    
    RedrawWindow m_hScrollBar, ByVal 0, ByVal 0, RDW_INVALIDATE Or RDW_UPDATENOW
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageUp Then
        SendMessage UserControl.hwnd, WM_VSCROLL, SB_PAGEUP, ByVal 0&
    ElseIf KeyCode = vbKeyPageDown Then
        SendMessage UserControl.hwnd, WM_VSCROLL, SB_PAGEDOWN, ByVal 0&
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbKeyLButton Then
        If m_wasOverUrl Then
            RaiseEvent clickedUrl(m_url)
        Else
            m_selecting = True
            m_selectStartX = x
            m_selectStartY = y
            
            m_realSelectStartX = x
            m_realSelectStartY = y
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo UserControl_MouseMove_Error

    m_mouseX = x
    m_mouseY = y
    
    If m_selecting Then
        UserControl.MousePointer = vbIbeam
    
        If x > m_realSelectStartX Then
            If x < 0 Then
                x = 0
            Else
                m_selectEndX = x
            End If
        
            m_selectStartX = m_realSelectStartX
        Else
            m_selectEndX = m_realSelectStartX
            
            If x < 0 Then
                m_selectStartX = 0
            Else
                m_selectStartX = x
            End If
        End If
        
        If y > m_realSelectStartY Then
            m_selectEndY = y
            m_selectStartY = m_realSelectStartY
        Else
            m_selectEndY = m_realSelectStartY
            m_selectStartY = y
        End If
        
        If x < 0 Then
            m_moveX = 0
        Else
            m_moveX = x
        End If
        
        processSelection
    Else
        If y < 0 Or y > UserControl.ScaleHeight Then
            If m_wasOverUrl Then
                UserControl.MousePointer = 0
                m_wasOverUrl = False
            End If
            
            Exit Sub
        End If
        
        Dim physLine As CPhysLine
        Set physLine = getLineByCoords(y)
        
        If physLine Is Nothing Then
            m_wasOverUrl = False
            Exit Sub
        End If
        
        Dim block As ITextRenderBlock
        Set block = physLine.getMouseOverBlock(x)

        Dim isUrl As Boolean

        If Not block Is Nothing Then
            If TypeOf block Is CBlockText Then
                Dim textBlock As CBlockText
            
                Set textBlock = block
                
                If textBlock.isUrl Then
                    If Not m_wasOverUrl Then
                        If Not g_handCursor Is Nothing Then
                            UserControl.MouseIcon = g_handCursor
                            UserControl.MousePointer = vbCustom
                        End If
                        
                        m_url = textBlock.url
                        m_wasOverUrl = True
                    End If
                    
                    isUrl = True
                End If
            End If
        End If
        
        If Not isUrl Then
            If m_wasOverUrl Then
                UserControl.MousePointer = 0
                m_wasOverUrl = False
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

UserControl_MouseMove_Error:
    handleError "UserControl_MouseMove", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Function getLineByCoords(ByVal y As Long) As CPhysLine
    Dim count As Long
    Dim linesUp As Long
    Dim pline As Long
    Dim vline As Long
    
    linesUp = Fix((UserControl.ScaleHeight - y) / m_fontHeight)
    
    vline = m_currentVirtLine
    pline = m_currentPhysLine
    
    For count = linesUp To 1 Step -1
        If pline = 1 Then
            vline = vline - 1
                            
            If vline < 1 Then
                Exit Function
            End If
            
            pline = m_lines.item(vline).physLineCount
        Else
            pline = pline - 1
        End If
    Next count
    
    If vline <= m_lines.count And m_lines.count > 0 Then
        If pline <= m_lines.item(vline).realPhysLineCount Then
            If m_lines.item(vline).realPhysLineCount > 0 Then
                Set getLineByCoords = m_lines.item(vline).physLine(pline)
            End If
        End If
    End If
End Function

Private Sub processSelection()
    Dim count As Integer

    For count = 1 To m_visibleLines.count
        If m_visibleLines.item(count).bottom >= m_selectStartY And m_visibleLines.item(count).top <= m_selectEndY Then
            m_visibleLines.item(count).setSelection m_selectStartY, m_selectEndY, m_selectStartX, m_selectEndX, m_realSelectStartX, m_realSelectStartY, m_moveX
        Else
            If m_visibleLines.item(count).selected Then
                m_visibleLines.item(count).unSelect
            End If
        End If
    Next count
    
    reDraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Button = vbKeyLButton Then
        Exit Sub
    End If

    m_selecting = False
    
    Dim count As Integer
    
    copySelectedText
    
    For count = 1 To m_visibleLines.count
        If m_visibleLines.item(count).selected Then
            m_visibleLines.item(count).unSelect
        End If
    Next count
    
    If m_wasOverUrl Then
        UserControl.MouseIcon = g_handCursor
        UserControl.MousePointer = vbCustom
    Else
        UserControl.MousePointer = vbDefault
    End If
    
    reDraw
    
    RaiseEvent noLongerNeedFocus
End Sub

Private Sub copySelectedText()
    Dim count As Integer
    
    Dim text As String
    
    For count = 1 To m_visibleLines.count
        If m_visibleLines.item(count).selected Then
            text = text & m_visibleLines.item(count).getSelectedText(m_drawingData, False)
        End If
    Next count
    
    If LenB(text) <> 0 Then
        Clipboard.clear
        Clipboard.SetText UTF8Encode(text)
    End If
End Sub

Private Sub UserControl_Paint()
    reDraw
End Sub

Private Sub drawLines(first As Long, y As Long, height As Integer)
    If m_fontmanager Is Nothing Then
        Exit Sub
    End If
    
    Dim oldFont As Long
    
    m_drawingData.left = 0
    m_drawingData.right = UserControl.ScaleWidth - SB_WIDTH
    m_drawingData.top = UserControl.ScaleHeight - (UserControl.ScaleHeight - y)
    m_drawingData.bottom = (y + height)
    
    m_drawingData.x = 0
    m_drawingData.y = m_drawingData.bottom - m_fontHeight

    m_drawingData.setPalette colourThemes.currentTheme.getPalette

    Dim count As Integer
    Dim line As CLine
    Dim srcY As Integer
    
    Dim realBottom As Integer
    Dim count2 As Integer
    
    realBottom = y + height
    
    m_drawingData.realY = y + height
    
    Dim found As Boolean
    Dim pos As Integer
    
    Dim newLines As New cArrayList
    
    For count = first To 1 Step -1
        m_drawingData.reset
        
        Set line = m_lines.item(count)

        If count = m_currentVirtLine Then
            wrap line, True
            line.render m_drawingData, m_currentPhysLine
        Else
            wrap line, False
            line.render m_drawingData, 0
        End If
        
        If Not line.visible Then
            newLines.Add line
            line.visible = True
        End If
        
        If m_drawingData.y <= (m_drawingData.top - m_drawingData.fontHeight) Then
            Exit For
        End If
    Next count
    
    If first = m_currentVirtLine Then
        For count = newLines.count To 1 Step -1
            m_visibleLines.Add newLines.item(count)
        Next count
    Else
        For count = 1 To newLines.count
            m_visibleLines.Add newLines.item(count), 1
        Next count
    End If
    
    m_drawingData.fillSpace
End Sub

Private Sub wrap(line As CLine, first As Boolean)
    If Not line.wrapped Then
        Dim diff As Integer
    
        diff = line.wordWrap(m_drawingData)

        If first = True Then
            If diff < 0 Then
                If m_currentPhysLine <> 1 Then
                    If m_currentPhysLine + diff < 1 Then
                        m_currentRealLine = m_currentRealLine - (m_currentPhysLine - 1)
                        m_currentPhysLine = 1
                    Else
                        m_currentPhysLine = m_currentPhysLine + diff
                        m_currentRealLine = m_currentRealLine + diff
                    End If
                End If
                
                m_realLineCount = m_realLineCount + diff
            Else
                m_realLineCount = m_realLineCount + diff
            End If
        Else
            m_realLineCount = m_realLineCount + diff
            m_currentRealLine = m_currentRealLine + diff
        End If
    End If
End Sub

Private Sub updateBackbuffer()
    If m_backBitmap <> 0 Then
        DeleteObject m_backBitmap
    End If
    
    If m_backBuffer <> 0 Then
        DeleteDC m_backBuffer
    End If
    
    m_backBuffer = CreateCompatibleDC(UserControl.hdc)
    m_backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth - SB_WIDTH, UserControl.ScaleHeight)
        
    SelectObject m_backBuffer, m_backBitmap
    
    SetBkMode m_backBuffer, OPAQUE
    
    If Not m_fontmanager Is Nothing Then
        SelectObject m_backBuffer, m_fontmanager.getFont(False, False, False)
    End If
    
    m_drawingData.Dc = m_backBuffer
    m_drawingData.width = UserControl.ScaleWidth - SB_WIDTH
End Sub

Private Sub displayBackBuffer()
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth - SB_WIDTH, UserControl.ScaleHeight, m_backBuffer, 0, 0, vbSrcCopy
End Sub

Private Sub reDraw()
    If m_fontmanager Is Nothing Then
        Exit Sub
    End If

    Dim count As Integer
    
    For count = 1 To m_visibleLines.count
        m_visibleLines.item(count).visible = False
    Next count
    
    m_visibleLines.clear

    Dim line As CLine

    If m_lines.count <> 0 Then
        If m_atBottom Then
            Set line = m_lines.item(m_currentVirtLine)
            wrap line, True
            
            m_currentRealLine = m_currentRealLine + (line.physLineCount - m_currentPhysLine)
            m_currentPhysLine = line.physLineCount
        End If
    
        drawLines CInt(m_currentVirtLine), 0, UserControl.ScaleHeight
        displayBackBuffer
    Else
        m_drawingData.top = 0
        m_drawingData.bottom = UserControl.ScaleHeight
        m_drawingData.left = 0
        m_drawingData.right = UserControl.ScaleWidth - SB_WIDTH
        m_drawingData.y = UserControl.ScaleHeight
        m_drawingData.fillSpace
        displayBackBuffer
    End If
End Sub

Private Sub hardRedraw()
    Dim count As Long
    
    For count = 1 To m_visibleLines.count
        m_visibleLines.item(count).needsWrapping
    Next count
    
    reDraw
End Sub

Private Sub UserControl_Resize()
    MoveWindow m_hScrollBar, UserControl.ScaleWidth - SB_WIDTH, 0, SB_WIDTH, UserControl.ScaleHeight, False

    m_pageLines = Fix(UserControl.ScaleHeight / m_fontHeight)
    updateScrollBar
    
    Dim count As Integer
    Dim diff As Long
    
    If m_lines.count > 0 Then
        For count = m_lines.count To 1 Step -1
            If m_lines.item(count).wrapped = False Then
                Exit For
            End If
            
            m_lines.item(count).clearWrap
        Next count
        
        If count > 1 Then
            For count = 1 To m_lines.count
                If m_lines.item(count).wrapped = False Then
                    Exit For
                End If
                
                m_lines.item(count).clearWrap
            Next count
        End If
        
        For count = m_currentVirtLine To 1 Step -1
            If m_lines.item(count).wrapped = False Then
                Exit For
            End If
            
            m_lines.item(count).clearWrap
        Next count
    End If
    
    updateBackbuffer
    
    reDraw
    updateScrollBar
End Sub

Public Sub refresh()
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    deInitMessages
    closeLog
End Sub

Public Sub writeEntireBuffer()
    Dim count As Long
    Dim text As String
    Dim numberWritten As Long
    Dim codes As Boolean
    Dim timestamps As Boolean
    
    codes = settings.setting("logIncludeCodes", estBoolean)
    timestamps = settings.setting("logIncludeTimestamp", estBoolean)
    
    For count = 1 To m_lines.count
        text = m_lines.item(count).getText(codes, timestamps, m_drawingData.ignoreSeperators)
        logRealWrite text
    Next count
End Sub

Private Sub openLog()
    m_logFilename = m_logBaseName & "." & getFileNameDate & ".log"
    m_logHandle = CreateFile(StrPtr(m_logFilename), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, ByVal 0)

    If m_logHandle = -1 Then
        Exit Sub
    End If

    SetFilePointer m_logHandle, 0, ByVal 0, FILE_END
    
    Dim fileSize1 As Long
    Dim fileSize2 As Long
    
    Dim numberWritten As Long
    
    fileSize1 = GetFileSize(m_logHandle, fileSize2)
    
    If fileSize1 = 0 And fileSize2 = 0 Then
        Dim bom As String
        
        bom = ChrW$(&HFEFF)
        WriteFile m_logHandle, ByVal StrPtr(bom), LenB(bom), numberWritten, ByVal 0
    End If
    
    Dim text As String
    
    text = vbCrLf & "Log started: " & formatTime(getSystemTime()) & vbCrLf
    WriteFile m_logHandle, ByVal StrPtr(text), LenB(text), numberWritten, ByVal 0
End Sub

Private Sub logWrite(line As CLine)
    Dim text As String
    
   On Error GoTo logWrite_Error

    text = line.getText(settings.setting("logIncludeCodes", estBoolean), settings.setting("logIncludeTimestamp", estBoolean), m_drawingData.ignoreSeperators)
    logRealWrite text

   On Error GoTo 0
   Exit Sub

logWrite_Error:
    handleError "logWrite", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub logRealWrite(text As String)
   On Error GoTo logRealWrite_Error

    If Not m_enableLogging Then
        Exit Sub
    End If

    If m_logHandle = -1 Then
        openLog
    Else
        If m_logFilename <> m_logBaseName & "." & getFileNameDate & ".log" Then
            closeLog
            openLog
        End If
    End If

    SetFilePointer m_logHandle, 0, ByVal 0, FILE_END

    Dim numberWritten As Long
    WriteFile m_logHandle, ByVal StrPtr(text), LenB(text), numberWritten, ByVal 0

   On Error GoTo 0
   Exit Sub

logRealWrite_Error:
    handleError "logRealWrite", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub closeLog()
    If m_logHandle <> -1 Then
        Dim text As String
        Dim numberWritten As Long
        
        text = vbCrLf & "Log ended: " & formatTime(getSystemTime()) & vbCrLf
    
        WriteFile m_logHandle, ByVal StrPtr(text), LenB(text), numberWritten, ByVal 0
        CloseHandle m_logHandle
    End If
    
    m_logHandle = -1
End Sub
