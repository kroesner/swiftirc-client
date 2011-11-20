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

Private m_fontManager As CFontManager

Private Sub scrolled()
10        If m_currentRealLine >= m_realLineCount Then
20            m_atBottom = True
30        Else
40            m_atBottom = False
50        End If
          
60        If m_wasOverUrl Then
70            UserControl.MousePointer = 0
80            m_wasOverUrl = False
90        End If
End Sub

Public Property Get logName() As String
10        logName = m_logBaseName
End Property

Public Property Let logName(newValue As String)
10        m_logBaseName = newValue
End Property

Public Property Get enableLogging() As Boolean
10        enableLogging = m_enableLogging
End Property

Public Property Let enableLogging(newValue As Boolean)
10        m_enableLogging = newValue
End Property

Public Property Get ignoreSeperators() As Boolean
10        ignoreSeperators = m_drawingData.ignoreSeperators
End Property

Public Property Let ignoreSeperators(newValue As Boolean)
10        m_drawingData.ignoreSeperators = newValue
End Property

Public Sub clear()
10        m_currentRealLine = 1
20        m_currentVirtLine = 1
30        m_currentPhysLine = 1
40        m_realLineCount = 1

50        scrolled

60        m_visibleLines.clear
70        m_lines.clear
80        m_allLines.clear
90        refresh
End Sub

Private Sub updateColours()
10        m_drawingData.defaultForeColour = g_textViewFore
20        m_drawingData.defaultBackColour = g_textViewBack
30        reDraw
End Sub

Public Property Let eventManager(newValue As CEventManager)
10        Set m_eventManager = newValue
End Property

Public Property Get foreColour() As Byte
10        foreColour = m_drawingData.defaultForeColour
End Property

Public Property Let foreColour(newValue As Byte)
10        m_drawingData.defaultForeColour = newValue
20        reDraw
End Property

Public Property Get backColour() As Byte
10        backColour = m_drawingData.defaultBackColour
End Property

Public Property Let backColour(newValue As Byte)
10        m_drawingData.defaultBackColour = newValue
20        reDraw
End Property

Public Sub addEvent(eventName As String, params() As String)
10        addEventEx eventName, Nothing, vbNullString, TVE_NONE, params
End Sub

Public Sub addEventEx(eventName As String, userStyle As CUserStyle, username As String, flags As _
    Long, params() As String)
          Dim line As New CLine
          Dim aEvent As CEvent
          
10       On Error GoTo addEventEx_Error

20        Set aEvent = m_eventManager.findEvent(eventName)
          
30        If aEvent Is Nothing Then
40            Exit Sub
50        End If
          
60        line.init aEvent, flags, userStyle, username, params
          
70        m_allLines.Add line
          
80        If Not line.shouldShow Then
90            Exit Sub
100       End If
          
110       If m_lines.count > 0 Then
120           If m_lines.item(m_lines.count).seperatorBottom(m_drawingData.ignoreSeperators) Then
130               line.seperatorAbove = True
140           End If
150       End If
          
160       m_lines.Add line
          
170       If m_atBottom Then
180           m_realLineCount = m_realLineCount + 1
              
190           If UserControl.Extender.visible = True Then
200               wrap line, True
210               scrollDown m_realLineCount - m_currentRealLine
220           Else
230               If m_lines.count > 1 Then
240                   m_currentRealLine = m_currentRealLine + 1
250                   m_currentVirtLine = m_lines.count
260                   m_currentPhysLine = 1
270               End If
280           End If
290       Else
300           m_realLineCount = m_realLineCount + 1
310       End If
          
320       line.wasDisplayed = True
330       removeOldLines
          
340       If m_realLineCount = 1 Then
350           reDraw
360       End If
          
370       updateScrollBar
          
380       logWrite line
          
390      On Error GoTo 0
400      Exit Sub

addEventEx_Error:
410       handleError "addEventEx", Err.Number, Err.Description, Erl, eventName
End Sub

Public Sub addRawText(format As String, params() As String)
10        addRawTextEx eventColours.otherText, 0, format, Nothing, vbNullString, TVE_NONE, params
End Sub

Public Sub addRawTextEx(eventColour As CEventColour, foreColour As Byte, format As String, _
    userStyle As CUserStyle, username As String, flags As Long, params() As String)
          
          Dim line As New CLine
          
10        On Error GoTo addRawTextEx_Error

20        line.initEx eventColour, foreColour, format, userStyle, username, flags Or TVE_NOEVENT Or _
              TVE_VISIBLE, params
          
30        m_allLines.Add line
40        m_lines.Add line
          
50        If m_atBottom Then
60            If UserControl.Extender.visible Then
70                m_realLineCount = m_realLineCount + 1
80                wrap line, True
90                scrollDown m_realLineCount - m_currentRealLine
100           Else
110               If m_lines.count > 1 Then
120                   m_currentRealLine = m_currentRealLine + 1
130                   m_currentVirtLine = m_lines.count
140                   m_currentPhysLine = 1
150                   m_realLineCount = m_realLineCount + 1
160               End If
170           End If
180       Else
190           m_realLineCount = m_realLineCount + 1
200       End If
          
210       line.wasDisplayed = True

220       If m_realLineCount = 1 Then
230           reDraw
240       End If

250       removeOldLines
260       updateScrollBar
          
270       logWrite line

280      On Error GoTo 0
290      Exit Sub

addRawTextEx_Error:
300       handleError "addRawTextEx", Err.Number, Err.Description, Erl, format
End Sub

Private Sub removeOldLines()
10        If m_lines.count <= MAX_LINES Then
20            Exit Sub
30        End If

          Dim lineCount As Long
          
40        lineCount = m_lines.count - MAX_LINES
          
50        If lineCount > MAX_REMOVE_LINES Then
60            lineCount = MAX_REMOVE_LINES
70        End If

80        If m_currentVirtLine <= lineCount Or m_currentRealLine <= m_pageLines + MAX_REMOVE_LINES Then
90            Exit Sub
100       End If
          
110       If lineCount = 0 Then
120           Exit Sub
130       End If
          
          Dim count As Long
          
140       For count = 1 To lineCount
150           m_realLineCount = m_realLineCount - m_lines.item(1).physLineCount
160           m_currentRealLine = m_currentRealLine - m_lines.item(1).physLineCount
170           m_currentVirtLine = m_currentVirtLine - 1
180           m_lines.Remove 1
190       Next count
          
          Dim removed As Long
          
200       For count = 1 To m_allLines.count
210           If removed = lineCount Then
220               Exit For
230           End If
              
240           If m_allLines.item(1).wasDisplayed Then
250               removed = removed + 1
260           End If
              
270           m_allLines.Remove 1
280       Next count
End Sub

Public Sub updateVisibility()
10        calculateEventVisibility
End Sub

Private Sub calculateEventVisibility()
          Dim count As Integer
          Dim line As CLine
          Dim origVirtLine As Integer
          
10        origVirtLine = m_currentVirtLine
          
20        For count = 1 To m_lines.count
30            Set line = m_lines.item(count)
          
40            If Not line.shouldShow Then
50                m_realLineCount = m_realLineCount - line.physLineCount
                  
60                If count <= origVirtLine Then
70                    m_currentRealLine = m_currentRealLine - line.physLineCount
80                    m_currentVirtLine = m_currentVirtLine - 1
90                End If
                  
100               If line.wrapped Then
110                   line.clearWrap
120               End If
                  
130               line.wasDisplayed = False
140           End If
150       Next count
          
160       m_lines.clear
          
          Dim linesAdded As Boolean
          
170       For count = 1 To m_allLines.count
180           Set line = m_allLines.item(count)
              
190           If line.shouldShow Then
200               If Not line.wasDisplayed Then
210                   linesAdded = True
                  
220                   m_realLineCount = m_realLineCount + line.physLineCount
                      
230                   If count <= m_currentVirtLine Then
240                       m_currentRealLine = m_currentRealLine + line.physLineCount
250                       m_currentVirtLine = m_currentVirtLine + 1
260                   End If
                      
270                   line.wasDisplayed = True
280               End If
                  
290               m_lines.Add line
300           End If
310       Next count
          
320       If linesAdded Then
330           If m_currentVirtLine = 0 Then
340               m_currentVirtLine = 1
350               m_currentPhysLine = 1
360               m_currentRealLine = 1
370           End If
380       End If
          
390       If m_lines.count > 0 Then
400           If m_currentPhysLine > m_lines.item(m_currentVirtLine).physLineCount Then
410               m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount
420           End If
430       End If
          
440       refresh
End Sub

Private Sub IColourUser_coloursUpdated()
10        m_drawingData.setPalette colourThemes.currentTheme.getPalette
20        m_drawingData.defaultBackColour = g_textViewBack
30        m_drawingData.defaultForeColour = g_textViewFore
40        hardRedraw
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
10        m_fontHeight = m_fontManager.fontHeight
20        m_drawingData.fontHeight = m_fontHeight
30        m_drawingData.fontManager = m_fontManager
40        refresh
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        Select Case CurrentMessage
              Case WM_SYSCOLORCHANGE
20                ISubclass_MsgResponse = emrPreprocess
30            Case Else
40                ISubclass_MsgResponse = emrConsume
50        End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
10        Select Case iMsg
              Case WM_VSCROLL
                  Dim scrollCode As Integer
                  
20                scrollCode = LoWord(wParam)
30                processScroll scrollCode
40            Case WM_MOUSEWHEEL
                  Dim accumDelta As Integer
                  
50                accumDelta = HiWord(wParam)
60                processMouseWheel accumDelta
70        End Select
End Function

Public Sub processMouseWheel(ByVal accumDelta As Integer)
          Dim lines As Long
          
10        Do While accumDelta >= 40
20            lines = lines + 1
30            accumDelta = accumDelta - 40
40        Loop
          
50        If lines > 0 Then
60            If lines >= m_currentRealLine Then
70                lines = m_currentRealLine - 1
80            End If
              
90            scrollUp lines
100       End If
          
110       lines = 0
          
120       Do While accumDelta <= -40
130           lines = lines + 1
140           accumDelta = accumDelta + 40
150       Loop
          
160       If lines > 0 Then
170           If lines > m_realLineCount - m_currentRealLine Then
180               lines = m_realLineCount - m_currentRealLine
190           End If
              
200           scrollDown lines
210       End If
          
220       updateScrollBar
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub UserControl_DblClick()
10        RaiseEvent doubleClick
End Sub

Private Sub UserControl_Initialize()
10        m_logHandle = -1

20        m_fontHeight = 20

30        m_drawingData.setPalette colourThemes.currentTheme.getPalette
40        m_drawingData.fontHeight = m_fontHeight
50        m_drawingData.fontManager = m_fontManager

60        initScrollBar
70        initMessages

80        m_currentRealLine = 1
90        m_currentVirtLine = 1
100       m_currentPhysLine = 1
          
110       m_atBottom = True
          
120       updateScrollBar
          
130       updateColours
End Sub

Private Sub initMessages()
10        AttachMessage Me, UserControl.hwnd, WM_CTLCOLORSCROLLBAR
20        AttachMessage Me, UserControl.hwnd, WM_VSCROLL
30        AttachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
40        AttachMessage Me, UserControl.hwnd, WM_SYSCOLORCHANGE
End Sub

Private Sub deInitMessages()
10        DetachMessage Me, UserControl.hwnd, WM_CTLCOLORSCROLLBAR
20        DetachMessage Me, UserControl.hwnd, WM_VSCROLL
30        DetachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
40        DetachMessage Me, UserControl.hwnd, WM_SYSCOLORCHANGE
End Sub

Private Sub scrollTop()
10        m_currentVirtLine = 1
20        m_currentPhysLine = 1
30        m_currentRealLine = 1
          
40        scrolled
          
50        updateScrollBar
60        reDraw
End Sub

Private Sub scrollBottom()
          Dim line As CLine
          
10        m_currentVirtLine = m_lines.count
          
20        Set line = m_lines.item(m_currentVirtLine)
          
30        wrap line, True
          
40        m_currentPhysLine = line.physLineCount
50        m_currentRealLine = m_realLineCount
          
60        scrolled
          
70        updateScrollBar
80        reDraw
End Sub

Private Sub scrollUp(lines As Long)
10       On Error GoTo scrollUp_Error

20        If lines < 1 Then
30            Exit Sub
40        End If
          
          Dim count As Long
          Dim lastVisibleLine As Long
          Dim visibleLines As Long
          Dim line As CLine
          
50        For count = m_currentVirtLine To 1 Step -1
60            Set line = m_lines.item(count)
         
70            If count = m_currentVirtLine Then
80                visibleLines = visibleLines + m_currentPhysLine
90            Else
100               visibleLines = visibleLines + line.physLineCount
110           End If
         
120           If visibleLines > m_pageLines Then
130               If count < m_currentVirtLine Then
140                   visibleLines = visibleLines - line.physLineCount
150               End If
                  
160               Exit For
170           End If
180       Next count
          
190       lastVisibleLine = count
          
200       For count = lines To 1 Step -1
210           If m_currentPhysLine > 1 Then
220               m_currentPhysLine = m_currentPhysLine - 1
230           Else
240               If m_currentVirtLine = 1 Then
250                   lines = lines - count
260                   Exit For
270               End If
                  
280               If m_visibleLines.count > 0 Then
290                   If m_visibleLines.item(m_visibleLines.count).selected Then
300                       m_visibleLines.item(m_visibleLines.count).unSelect
310                   End If
                  
320                   m_visibleLines.item(m_visibleLines.count).clearWrap
330                   m_visibleLines.item(m_visibleLines.count).visible = False
                      
340                   m_visibleLines.Remove m_visibleLines.count
350               End If
                  
360               m_currentVirtLine = m_currentVirtLine - 1
370               wrap m_lines.item(m_currentVirtLine), True
380               m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount
390           End If
400       Next count
          
410       If lines < 1 Then
420           Exit Sub
430       End If
          
440       m_currentRealLine = m_currentRealLine - lines
          
450       If lines >= m_pageLines Then
460           reDraw
470           Exit Sub
480       End If

          Dim copyHeight As Long
          Dim topMargin As Long
          Dim copyY As Long
          
490       topMargin = UserControl.ScaleHeight - (visibleLines * m_fontHeight)
500       copyHeight = m_fontHeight * (visibleLines - lines)
510       copyY = topMargin + (m_fontHeight * lines)
          
520       BitBlt m_backBuffer, 0, copyY, UserControl.ScaleWidth - SB_WIDTH, copyHeight, m_backBuffer, 0, topMargin, _
              vbSrcCopy
          
530       drawLines lastVisibleLine, 0, CInt(copyY)
540       displayBackBuffer
          
550       scrolled

560      On Error GoTo 0
570      Exit Sub

scrollUp_Error:
580       handleError "scrollUp", Err.Number, Err.Description, Erl, CStr(lines) & " lines"
End Sub

Private Sub scrollDown(lines As Long)
10       On Error GoTo scrollDown_Error

20        If lines < 1 Then
30            Exit Sub
40        End If
          
          Dim count As Long
          Dim diff As Long
          
50        For count = lines To 1 Step -1
60            If m_lines.item(m_currentVirtLine).physLineCount > m_currentPhysLine Then
70                m_currentPhysLine = m_currentPhysLine + 1
80            Else
90                If m_lines.count <= m_currentVirtLine Then
100                   lines = lines - count
110                   Exit For
120               End If
                  
130               m_currentVirtLine = m_currentVirtLine + 1
140               m_currentPhysLine = 1
                  
150               wrap m_lines.item(m_currentVirtLine), True
160           End If
170       Next count
          
180       If lines < 1 Then
190           Exit Sub
200       End If
          
210       m_currentRealLine = m_currentRealLine + lines
          
220       If lines >= m_pageLines Then
230           reDraw
240           Exit Sub
250       End If
          
          Dim scrollDist As Long
          Dim copyHeight As Long
          
260       scrollDist = m_fontHeight * lines
270       copyHeight = UserControl.ScaleHeight - scrollDist
          
280       BitBlt m_backBuffer, 0, 0, UserControl.ScaleWidth - SB_WIDTH, copyHeight, m_backBuffer, 0, scrollDist, _
              vbSrcCopy
          
290       For count = 1 To m_visibleLines.count
300           m_visibleLines.item(count).shiftedUp scrollDist
310       Next count
          
320       Do While m_visibleLines.count > 0
330           If m_visibleLines.item(1).bottom >= 0 Then
340               Exit Do
350           End If
              
360           If m_visibleLines.item(1).selected Then
370               m_visibleLines.item(1).unSelect
380           End If
              
390           m_visibleLines.item(1).visible = False
400           m_visibleLines.item(1).clearWrap
              
410           m_visibleLines.Remove 1
420       Loop
          
430       drawLines m_currentVirtLine, copyHeight, UserControl.ScaleHeight - copyHeight
440       displayBackBuffer
          
450       scrolled

460      On Error GoTo 0
470      Exit Sub

scrollDown_Error:
480       handleError "scrollDown", Err.Number, Err.Description, Erl, CStr(lines) & " lines"
End Sub

Public Sub pageUp()
10        If m_currentRealLine - m_pageLines < 1 Then
20            scrollUp m_currentRealLine - 1
30        Else
40            scrollUp m_pageLines
50        End If
          
60        scrolled
70        updateScrollBar
End Sub

Public Sub pageDown()
10        If m_currentRealLine + m_pageLines > m_realLineCount Then
20            scrollDown m_realLineCount - m_currentRealLine
30        Else
40            scrollDown m_pageLines
50        End If
          
60        scrolled
70        updateScrollBar
End Sub

Private Sub processScroll(scrollCode As Integer)
10       On Error GoTo processScroll_Error

20        Select Case scrollCode
              Case SB_PAGEUP
30                pageUp
40            Case SB_PAGEDOWN
50                pageDown
60            Case SB_LINEUP
70                If m_currentRealLine > 1 Then
80                    scrollUp 1
90                    updateScrollBar
100               End If
110           Case SB_LINEDOWN
120               If m_currentRealLine < m_realLineCount Then
130                   scrollDown 1
140                   updateScrollBar
150               End If
160           Case SB_TOP
170               scrollTop
180           Case SB_BOTTOM
190               scrollBottom
200           Case SB_THUMBTRACK
                  Dim si As SCROLLINFO
                  Dim diff As Long
                  
210               si.cbSize = Len(si)
220               si.fMask = SIF_TRACKPOS
                  
230               GetScrollInfo m_hScrollBar, SB_CTL, si
                  
240               If si.nTrackPos = 1 Then
250                   scrollTop
260                   Exit Sub
270               ElseIf si.nTrackPos >= m_realLineCount Then
280                   scrollBottom
290                   Exit Sub
300               End If
                  
                  Dim pos As Single
                  Dim vline As Integer
                  
310               diff = m_currentRealLine - si.nTrackPos
                  
320               If diff < 0 Then
330                   If -diff > (m_pageLines) Then
340                       bigScroll diff
350                   Else
360                       If m_currentRealLine - diff > m_realLineCount Then
370                           scrollDown m_realLineCount - m_currentRealLine
380                           updateScrollBar
390                       Else
400                           scrollDown -diff
410                           updateScrollBar
420                       End If
430                   End If
440               Else
450                   If diff > (m_pageLines) Then
460                       bigScroll diff
470                   Else
480                       If m_currentRealLine - diff < 1 Then
490                           scrollUp diff - m_currentRealLine
500                           updateScrollBar
510                       Else
520                           scrollUp diff
530                           updateScrollBar
540                       End If
550                   End If
560               End If
570       End Select

580      On Error GoTo 0
590      Exit Sub

processScroll_Error:
600       handleError "processScroll", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub bigScroll(diff As Long)
          Dim count As Long

10        For count = m_currentVirtLine To 1 Step -1
20            If m_lines.item(count).wrapped = False Then
30                Exit For
40            End If
                  
50            m_lines.item(count).clearWrap
60        Next count

70        If diff > 0 Then
80            For count = diff To 1 Step -1
90                If m_currentPhysLine = 1 Then
100                   If m_currentVirtLine = 1 Then Exit For
                      
110                   m_currentVirtLine = m_currentVirtLine - 1
120                   m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount
130               Else
140                   m_currentPhysLine = m_currentPhysLine - 1
150               End If
160           Next count
170       Else
180           For count = -diff To 1 Step -1
190               If m_currentPhysLine = m_lines.item(m_currentVirtLine).physLineCount Then
200                   If m_currentVirtLine = m_lines.count Then Exit For
                  
210                   m_currentVirtLine = m_currentVirtLine + 1
220                   m_currentPhysLine = 1
230               Else
240                   m_currentPhysLine = m_currentPhysLine + 1
250               End If
260           Next count
270       End If
          
280       m_currentRealLine = m_currentRealLine - diff
          
290       If m_currentRealLine < 1 Then
300           m_currentRealLine = 1
310       ElseIf m_currentRealLine > m_realLineCount Then
320           m_currentRealLine = m_realLineCount
330       End If
          
340       scrolled
          
350       updateScrollBar
360       reDraw
End Sub

Private Sub initScrollBar()
10        m_hScrollBar = CreateWindowEx(0, "SCROLLBAR", "", WS_CHILD Or SBS_VERT, UserControl.ScaleWidth _
              - SB_WIDTH, 0, SB_WIDTH, UserControl.ScaleHeight, UserControl.hwnd, 0, App.hInstance, ByVal 0&)
20        ShowScrollBar m_hScrollBar, SB_CTL, 1
          
30        m_scrollbarEnabled = True
          
40        updateScrollBar
End Sub

Private Sub updateScrollBar()
          Dim si As SCROLLINFO
          
10        si.cbSize = Len(si)
20        si.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
          
30        si.nMin = 1
          
40        If m_realLineCount < 2 Then
50            si.nMax = 0
60        Else
70            si.nMax = m_realLineCount + (m_pageLines - 1)
80        End If
          
90        si.nPage = m_pageLines
100       si.nPos = m_currentRealLine
          
110       SetScrollInfo m_hScrollBar, SB_CTL, si, 0
          
120       If si.nMax = 0 Then
130           EnableScrollBar m_hScrollBar, SB_CTL, ESB_DISABLE_BOTH
140       Else
150           EnableScrollBar m_hScrollBar, SB_CTL, ESB_ENABLE_BOTH
160       End If
          
170       RedrawWindow m_hScrollBar, ByVal 0, ByVal 0, RDW_INVALIDATE Or RDW_UPDATENOW
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeyPageUp Then
20            SendMessage UserControl.hwnd, WM_VSCROLL, SB_PAGEUP, ByVal 0&
30        ElseIf KeyCode = vbKeyPageDown Then
40            SendMessage UserControl.hwnd, WM_VSCROLL, SB_PAGEDOWN, ByVal 0&
50        End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If Button = vbKeyLButton Then
20            If m_wasOverUrl Then
30                RaiseEvent clickedUrl(m_url)
40            Else
50                m_selecting = True
60                m_selectStartX = x
70                m_selectStartY = y
                  
80                m_realSelectStartX = x
90                m_realSelectStartY = y
100           End If
110       End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
10       On Error GoTo UserControl_MouseMove_Error

20        m_mouseX = x
30        m_mouseY = y
          
40        If m_selecting Then
50            UserControl.MousePointer = vbIbeam
          
60            If x > m_realSelectStartX Then
70                If x < 0 Then
80                    x = 0
90                Else
100                   m_selectEndX = x
110               End If
              
120               m_selectStartX = m_realSelectStartX
130           Else
140               m_selectEndX = m_realSelectStartX
                  
150               If x < 0 Then
160                   m_selectStartX = 0
170               Else
180                   m_selectStartX = x
190               End If
200           End If
              
210           If y > m_realSelectStartY Then
220               m_selectEndY = y
230               m_selectStartY = m_realSelectStartY
240           Else
250               m_selectEndY = m_realSelectStartY
260               m_selectStartY = y
270           End If
              
280           If x < 0 Then
290               m_moveX = 0
300           Else
310               m_moveX = x
320           End If
              
330           processSelection
340       Else
350           If y < 0 Or y > UserControl.ScaleHeight Then
360               If m_wasOverUrl Then
370                   UserControl.MousePointer = 0
380                   m_wasOverUrl = False
390               End If
                  
400               Exit Sub
410           End If
              
              Dim physLine As CPhysLine
420           Set physLine = getLineByCoords(y)
              
430           If physLine Is Nothing Then
440               m_wasOverUrl = False
450               Exit Sub
460           End If
              
              Dim block As ITextRenderBlock
470           Set block = physLine.getMouseOverBlock(x)

              Dim isUrl As Boolean

480           If Not block Is Nothing Then
490               If TypeOf block Is CBlockText Then
                      Dim textBlock As CBlockText
                  
500                   Set textBlock = block
                      
510                   If textBlock.isUrl Then
520                       If Not m_wasOverUrl Then
530                           If Not g_handCursor Is Nothing Then
540                               UserControl.MouseIcon = g_handCursor
550                               UserControl.MousePointer = vbCustom
560                           End If
                              
570                           m_url = textBlock.url
580                           m_wasOverUrl = True
590                       End If
                          
600                       isUrl = True
610                   End If
620               End If
630           End If
              
640           If Not isUrl Then
650               If m_wasOverUrl Then
660                   UserControl.MousePointer = 0
670                   m_wasOverUrl = False
680               End If
690           End If
700       End If

710      On Error GoTo 0
720      Exit Sub

UserControl_MouseMove_Error:
730       handleError "UserControl_MouseMove", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Function getLineByCoords(ByVal y As Long) As CPhysLine
          Dim count As Long
          Dim linesUp As Long
          Dim pline As Long
          Dim vline As Long
          
10        linesUp = Fix((UserControl.ScaleHeight - y) / m_fontHeight)
          
20        vline = m_currentVirtLine
30        pline = m_currentPhysLine
          
40        For count = linesUp To 1 Step -1
50            If pline = 1 Then
60                vline = vline - 1
                                  
70                If vline < 1 Then
80                    Exit Function
90                End If
                  
100               pline = m_lines.item(vline).physLineCount
110           Else
120               pline = pline - 1
130           End If
140       Next count
          
150       If vline <= m_lines.count And m_lines.count > 0 Then
160           If pline <= m_lines.item(vline).realPhysLineCount Then
170               If m_lines.item(vline).realPhysLineCount > 0 Then
180                   Set getLineByCoords = m_lines.item(vline).physLine(pline)
190               End If
200           End If
210       End If
End Function

Private Sub processSelection()
          Dim count As Integer

10        For count = 1 To m_visibleLines.count
20            If m_visibleLines.item(count).bottom >= m_selectStartY And m_visibleLines.item(count).top _
                  <= m_selectEndY Then
30                m_visibleLines.item(count).setSelection m_selectStartY, m_selectEndY, m_selectStartX, _
                      m_selectEndX, m_realSelectStartX, m_realSelectStartY, m_moveX
40            Else
50                If m_visibleLines.item(count).selected Then
60                    m_visibleLines.item(count).unSelect
70                End If
80            End If
90        Next count
          
100       reDraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If Not Button = vbKeyLButton Then
20            Exit Sub
30        End If

40        m_selecting = False
          
          Dim count As Integer
          
50        copySelectedText
          
60        For count = 1 To m_visibleLines.count
70            If m_visibleLines.item(count).selected Then
80                m_visibleLines.item(count).unSelect
90            End If
100       Next count
          
110       If m_wasOverUrl Then
120           UserControl.MouseIcon = g_handCursor
130           UserControl.MousePointer = vbCustom
140       Else
150           UserControl.MousePointer = vbDefault
160       End If
          
170       reDraw
          
180       RaiseEvent noLongerNeedFocus
End Sub

Private Sub copySelectedText()
          Dim count As Integer
          
          Dim text As String
          
10        For count = 1 To m_visibleLines.count
20            If m_visibleLines.item(count).selected Then
30                text = text & m_visibleLines.item(count).getSelectedText(m_drawingData, False)
40            End If
50        Next count
          
60        If LenB(text) <> 0 Then
70            Clipboard.clear
80            Clipboard.SetText UTF8Encode(text)
90        End If
End Sub

Private Sub UserControl_Paint()
10        reDraw
End Sub

Private Sub drawLines(first As Long, y As Long, height As Integer)
10        If m_fontManager Is Nothing Then
20            Exit Sub
30        End If
          
          Dim oldFont As Long
          
40        m_drawingData.left = 0
50        m_drawingData.right = UserControl.ScaleWidth - SB_WIDTH
60        m_drawingData.top = UserControl.ScaleHeight - (UserControl.ScaleHeight - y)
70        m_drawingData.bottom = (y + height)
          
80        m_drawingData.x = 0
90        m_drawingData.y = m_drawingData.bottom - m_fontHeight

100       m_drawingData.setPalette colourThemes.currentTheme.getPalette

          Dim count As Integer
          Dim line As CLine
          Dim srcY As Integer
          
          Dim realBottom As Integer
          Dim count2 As Integer
          
110       realBottom = y + height
          
120       m_drawingData.realY = y + height
          
          Dim found As Boolean
          Dim pos As Integer
          
          Dim newLines As New cArrayList
          
130       For count = first To 1 Step -1
140           m_drawingData.reset
              
150           Set line = m_lines.item(count)

160           If count = m_currentVirtLine Then
170               wrap line, True
180               line.render m_drawingData, m_currentPhysLine
190           Else
200               wrap line, False
210               line.render m_drawingData, 0
220           End If
              
230           If Not line.visible Then
240               newLines.Add line
250               line.visible = True
260           End If
              
270           If m_drawingData.y <= (m_drawingData.top - m_drawingData.fontHeight) Then
280               Exit For
290           End If
300       Next count
          
310       If first = m_currentVirtLine Then
320           For count = newLines.count To 1 Step -1
330               m_visibleLines.Add newLines.item(count)
340           Next count
350       Else
360           For count = 1 To newLines.count
370               m_visibleLines.Add newLines.item(count), 1
380           Next count
390       End If
          
400       m_drawingData.fillSpace
End Sub

Private Sub wrap(line As CLine, first As Boolean)
10        If Not line.wrapped Then
              Dim diff As Integer
          
20            diff = line.wordWrap(m_drawingData)

30            If first = True Then
40                If diff < 0 Then
50                    If m_currentPhysLine <> 1 Then
60                        If m_currentPhysLine + diff < 1 Then
70                            m_currentRealLine = m_currentRealLine - (m_currentPhysLine - 1)
80                            m_currentPhysLine = 1
90                        Else
100                           m_currentPhysLine = m_currentPhysLine + diff
110                           m_currentRealLine = m_currentRealLine + diff
120                       End If
130                   End If
                      
140                   m_realLineCount = m_realLineCount + diff
150               Else
160                   m_realLineCount = m_realLineCount + diff
170               End If
180           Else
190               m_realLineCount = m_realLineCount + diff
200               m_currentRealLine = m_currentRealLine + diff
210           End If
220       End If
End Sub

Private Sub updateBackbuffer()
10        If m_backBitmap <> 0 Then
20            DeleteObject m_backBitmap
30        End If
          
40        If m_backBuffer <> 0 Then
50            DeleteDC m_backBuffer
60        End If
          
70        m_backBuffer = CreateCompatibleDC(UserControl.hdc)
80        m_backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth - SB_WIDTH, _
              UserControl.ScaleHeight)
              
90        SelectObject m_backBuffer, m_backBitmap
          
100       SetBkMode m_backBuffer, OPAQUE
          
110       If Not m_fontManager Is Nothing Then
120           SelectObject m_backBuffer, m_fontManager.getFont(False, False, False)
130       End If
          
140       m_drawingData.Dc = m_backBuffer
150       m_drawingData.width = UserControl.ScaleWidth - SB_WIDTH
End Sub

Private Sub displayBackBuffer()
10        BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth - SB_WIDTH, UserControl.ScaleHeight, _
              m_backBuffer, 0, 0, vbSrcCopy
End Sub

Private Sub reDraw()
10        If m_fontManager Is Nothing Then
20            Exit Sub
30        End If

          Dim count As Integer
          
40        For count = 1 To m_visibleLines.count
50            m_visibleLines.item(count).visible = False
60        Next count
          
70        m_visibleLines.clear

          Dim line As CLine

80        If m_lines.count <> 0 Then
90            If m_atBottom Then
100               Set line = m_lines.item(m_currentVirtLine)
110               wrap line, True
                  
120               m_currentRealLine = m_currentRealLine + (line.physLineCount - m_currentPhysLine)
130               m_currentPhysLine = line.physLineCount
140           End If
          
150           drawLines CInt(m_currentVirtLine), 0, UserControl.ScaleHeight
160           displayBackBuffer
170       Else
180           m_drawingData.top = 0
190           m_drawingData.bottom = UserControl.ScaleHeight
200           m_drawingData.left = 0
210           m_drawingData.right = UserControl.ScaleWidth - SB_WIDTH
220           m_drawingData.y = UserControl.ScaleHeight
230           m_drawingData.fillSpace
240           displayBackBuffer
250       End If
End Sub

Private Sub hardRedraw()
          Dim count As Long
          
10        For count = 1 To m_visibleLines.count
20            m_visibleLines.item(count).needsWrapping
30        Next count
          
40        reDraw
End Sub

Private Sub UserControl_Resize()
10        MoveWindow m_hScrollBar, UserControl.ScaleWidth - SB_WIDTH, 0, SB_WIDTH, UserControl.ScaleHeight, False

20        m_pageLines = Fix(UserControl.ScaleHeight / m_fontHeight)
30        updateScrollBar
          
          Dim count As Integer
          Dim diff As Long
          
40        If m_lines.count > 0 Then
50            For count = m_lines.count To 1 Step -1
60                If m_lines.item(count).wrapped = False Then
70                    Exit For
80                End If
                  
90                m_lines.item(count).clearWrap
100           Next count
              
110           If count > 1 Then
120               For count = 1 To m_lines.count
130                   If m_lines.item(count).wrapped = False Then
140                       Exit For
150                   End If
                      
160                   m_lines.item(count).clearWrap
170               Next count
180           End If
              
190           For count = m_currentVirtLine To 1 Step -1
200               If m_lines.item(count).wrapped = False Then
210                   Exit For
220               End If
                  
230               m_lines.item(count).clearWrap
240           Next count
250       End If
          
260       updateBackbuffer
          
270       reDraw
280       updateScrollBar
End Sub

Public Sub refresh()
10        UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
10        deInitMessages
20        closeLog
End Sub

Public Sub writeEntireBuffer()
          Dim count As Long
          Dim text As String
          Dim numberWritten As Long
          Dim codes As Boolean
          Dim timestamps As Boolean
          
10        codes = settings.setting("logIncludeCodes", estBoolean)
20        timestamps = settings.setting("logIncludeTimestamp", estBoolean)
          
30        For count = 1 To m_lines.count
40            text = m_lines.item(count).getText(codes, timestamps, m_drawingData.ignoreSeperators)
50            logRealWrite text
60        Next count
End Sub

Private Sub openLog()
10        m_logFilename = m_logBaseName & "." & getFileNameDate & ".log"
20        m_logHandle = CreateFile(StrPtr(m_logFilename), GENERIC_WRITE, _
              FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, ByVal 0)

30        If m_logHandle = -1 Then
40            Exit Sub
50        End If

60        SetFilePointer m_logHandle, 0, ByVal 0, FILE_END
          
          Dim fileSize1 As Long
          Dim fileSize2 As Long
          
          Dim numberWritten As Long
          
70        fileSize1 = GetFileSize(m_logHandle, fileSize2)
          
80        If fileSize1 = 0 And fileSize2 = 0 Then
              Dim bom As String
              
90            bom = ChrW$(&HFEFF)
100           WriteFile m_logHandle, ByVal StrPtr(bom), LenB(bom), numberWritten, ByVal 0
110       End If
          
          Dim text As String
          
120       text = vbCrLf & "Log started: " & formatTime(getSystemTime()) & vbCrLf
130       WriteFile m_logHandle, ByVal StrPtr(text), LenB(text), numberWritten, ByVal 0
End Sub

Private Sub logWrite(line As CLine)
          Dim text As String
          
10       On Error GoTo logWrite_Error

20        text = line.getText(settings.setting("logIncludeCodes", estBoolean), _
              settings.setting("logIncludeTimestamp", estBoolean), m_drawingData.ignoreSeperators)
30        logRealWrite text

40       On Error GoTo 0
50       Exit Sub

logWrite_Error:
60        handleError "logWrite", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub logRealWrite(text As String)
10       On Error GoTo logRealWrite_Error

20        If Not m_enableLogging Then
30            Exit Sub
40        End If

50        If m_logHandle = -1 Then
60            openLog
70        Else
80            If m_logFilename <> m_logBaseName & "." & getFileNameDate & ".log" Then
90                closeLog
100               openLog
110           End If
120       End If

130       SetFilePointer m_logHandle, 0, ByVal 0, FILE_END

          Dim numberWritten As Long
140       WriteFile m_logHandle, ByVal StrPtr(text), LenB(text), numberWritten, ByVal 0

150      On Error GoTo 0
160      Exit Sub

logRealWrite_Error:
170       handleError "logRealWrite", Err.Number, Err.Description, Erl, vbNullString
End Sub

Private Sub closeLog()
10        If m_logHandle <> -1 Then
              Dim text As String
              Dim numberWritten As Long
              
20            text = vbCrLf & "Log ended: " & formatTime(getSystemTime()) & vbCrLf
          
30            WriteFile m_logHandle, ByVal StrPtr(text), LenB(text), numberWritten, ByVal 0
40            CloseHandle m_logHandle
50        End If
          
60        m_logHandle = -1
End Sub
