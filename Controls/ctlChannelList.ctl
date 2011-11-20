VERSION 5.00
Begin VB.UserControl ctlChannelList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Menu mnuChannelListContext 
      Caption         =   "mnuChannelListContext"
      Begin VB.Menu mnuJoin 
         Caption         =   "Join channel"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Sort by"
         Begin VB.Menu mnuSortUsers 
            Caption         =   "Usercount"
         End
         Begin VB.Menu mnuSortName 
            Caption         =   "Name"
         End
         Begin VB.Menu mnuSortTopic 
            Caption         =   "Topic"
         End
      End
   End
End
Attribute VB_Name = "ctlChannelList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IColourUser
Implements ISubclass
Implements IFontUser
Implements ITabWindow
Implements IWindow

Private COLUMN_NAME As Long
Private COLUMN_USERS As Long
Private Const COLUMN_SPACING As Long = 10

Private Enum eSortType
    estName
    estUsers
    estTopic
End Enum

Private Enum eSortDirection
    esdAscending
    esdDescending
End Enum

Private m_listbox As Long

Private m_channels As New cArrayList

Private m_sortBy As eSortType

Private m_fontManager As CFontManager
Private m_realWindow As VBControlExtender

Private m_switchbarTab As CTab

Private m_session As CSession

Private m_fillbrush As Long
Private m_itemHeight As Long
Private m_maxTopicWidth As Long

Private m_labelManager As New CLabelManager

Public Event joinChannel(name As String)

Private m_IPAOHookStruct As IPAOHookStructChannelList

Public Property Get session() As CSession
10        Set session = m_session
End Property

Public Property Let session(newValue As CSession)
10        Set m_session = newValue
End Property

Public Property Get switchbartab() As CTab
10        Set switchbartab = m_switchbarTab
End Property

Public Property Let switchbartab(newValue As CTab)
10        Set m_switchbarTab = newValue
End Property

Private Sub IColourUser_coloursUpdated()

End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
10        fontsUpdated
End Sub

Private Sub fontsUpdated()
10        m_itemHeight = m_fontManager.fontHeight
20        UserControl_Resize
          
30        If Not m_listbox = 0 Then
40            SendMessage m_listbox, LB_SETITEMHEIGHT, 0, ByVal m_itemHeight
              
50            If m_channels.count > 0 Then
                  Dim TopIndex As Long
                  
60                TopIndex = SendMessage(m_listbox, LB_GETTOPINDEX, 0, ByVal 0)
                  
70                If TopIndex > (m_channels.count - 1) - (UserControl.ScaleHeight / m_itemHeight) Then
                      'If the font change has left space at the bottom of
                      'the listbox, update the scroll index appropriately.
80                    SendMessage m_listbox, LB_SETTOPINDEX, m_channels.count - 1, ByVal 0
90                End If
100           End If
110       End If
End Sub

Public Function TranslateAccelerator(lpMsg As VBOleGuids.Msg) As Long
10        TranslateAccelerator = S_FALSE
          ' Here you can modify the response to the key down
          ' accelerator command using the values in lpMsg.  This
          ' can be used to capture Tabs, Returns, Arrows etc.
          ' Just process the message as required and return S_OK.
          Dim key As Long
          
20        key = (lpMsg.wParam And &HFFFF&)
          
30        If key = vbKeyLeft Or key = vbKeyRight Or key = vbKeyUp Or key = vbKeyDown Then
40            If lpMsg.message = WM_KEYDOWN Or lpMsg.message = WM_KEYUP Then
50                SendMessage m_listbox, lpMsg.message, key, 0
60                TranslateAccelerator = S_OK
70            End If
80        End If
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        Select Case CurrentMessage
              Case WM_MOUSEWHEEL
20                ISubclass_MsgResponse = emrConsume
30            Case WM_LBUTTONDBLCLK
40                ISubclass_MsgResponse = emrPostProcess
50            Case Else
60                ISubclass_MsgResponse = emrPreprocess
70        End Select
End Property

Private Property Get ShiftState() As Integer
10        If GetAsyncKeyState(vbKeyShift) <> 0 Then
20            ShiftState = vbShiftMask
30        End If

40        If GetAsyncKeyState(vbKeyControl) <> 0 Then
50            ShiftState = ShiftState Or vbCtrlMask
60        End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
          Dim coordY As Long
          
10        Select Case iMsg
              Case WM_DRAWITEM
                  Dim item As DRAWITEMSTRUCT
                  
20                CopyMemory item, ByVal lParam, Len(item)
30                ISubclass_WindowProc = drawItem(item)
40            Case WM_MEASUREITEM
                  Dim measureItem As MEASUREITEMSTRUCT
                  
50                CopyMemory measureItem, ByVal lParam, Len(measureItem)
60                measureItem.itemHeight = m_itemHeight
70                CopyMemory ByVal lParam, measureItem, Len(measureItem)
              
80                ISubclass_WindowProc = True
90            Case WM_SETFOCUS
                  Dim pOleObject                  As IOleObject
                  Dim pOleInPlaceSite             As IOleInPlaceSite
                  Dim pOleInPlaceFrame            As IOleInPlaceFrame
                  Dim pOleInPlaceUIWindow         As IOleInPlaceUIWindow
                  Dim pOleInPlaceActiveObject     As IOleInPlaceActiveObject
                  Dim PosRect                     As RECT
                  Dim ClipRect                    As RECT
                  Dim FrameInfo                   As OLEINPLACEFRAMEINFO
                  Dim grfModifiers                As Long
                  Dim AcceleratorMsg              As Msg
                  
                  'Get in-place frame and make sure it is set to our in-between
                  'implementation of IOleInPlaceActiveObject in order to catch
                  'TranslateAccelerator calls
100               Set pOleObject = Me
110               Set pOleInPlaceSite = pOleObject.GetClientSite
120               pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), _
                      VarPtr(ClipRect), VarPtr(FrameInfo)
130               CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
140               pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
150               If Not pOleInPlaceUIWindow Is Nothing Then
160                  pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
170               End If
                  ' Clear up the inbetween implementation:
180               CopyMemory pOleInPlaceActiveObject, 0&, 4
                  ' --------------------------------------------------------------------------
190           Case WM_CTLCOLORLISTBOX
200               If m_fillbrush <> 0 Then
210                   DeleteObject m_fillbrush
220               End If
                  
230               m_fillbrush = CreateSolidBrush(getPaletteEntry(g_channelListBack))
240               ISubclass_WindowProc = m_fillbrush
250           Case WM_MOUSEWHEEL
260               If Not m_listbox = 0 Then
                      Dim accumDelta As Integer
          
270                   accumDelta = HiWord(wParam)
              
280                   Do While accumDelta >= 40
290                       SendMessage m_listbox, WM_VSCROLL, SB_LINEUP, 0
300                       accumDelta = accumDelta - 40
310                   Loop
              
320                   Do While accumDelta <= -40
330                       SendMessage m_listbox, WM_VSCROLL, SB_LINEDOWN, 0
340                       accumDelta = accumDelta + 40
350                   Loop
360               End If
370           Case WM_LBUTTONDBLCLK
                  Dim index As Long
                  
380               index = SendMessage(m_listbox, LB_GETCURSEL, 0, ByVal 0)
                  
390               If index <> LB_ERR Then
400                   RaiseEvent joinChannel(m_channels.item(index + 1).name)
410               End If
420           Case WM_RBUTTONDOWN
430               coordY = HiWord(lParam)
                  
                  Dim TopIndex As Long
                  Dim itemIndex As Long
                  
440               TopIndex = SendMessage(m_listbox, LB_GETTOPINDEX, 0, ByVal 0)
450               itemIndex = TopIndex + Fix(coordY / m_itemHeight)
                  
460               SendMessage m_listbox, LB_SETCURSEL, itemIndex, ByVal 0
470               PopupMenu mnuChannelListContext
480       End Select
End Function

Private Function drawItem(item As DRAWITEMSTRUCT) As Long
10        If item.itemID = -1 Then
20            Exit Function
30        End If
          
          Dim channel As CChannelListItem
          Dim selected As Boolean

40        Set channel = m_channels.item(item.itemID + 1)
          
50        selected = item.itemState And ODS_SELECTED
          
60        If selected Then
70            SetBkColor item.hdc, GetSysColor(COLOR_HIGHLIGHT)
80            SetTextColor item.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
90            FillRect item.hdc, item.rcItem, GetSysColorBrush(COLOR_HIGHLIGHT)
100       Else
110           SetBkColor item.hdc, getPaletteEntry(g_channelListBack)
120           SetTextColor item.hdc, getPaletteEntry(g_channelListFore)
              
              Dim fillBrush As Long
130           fillBrush = CreateSolidBrush(getPaletteEntry(g_channelListBack))
140           FillRect item.hdc, item.rcItem, fillBrush
150           DeleteObject fillBrush
160       End If
          
170       SelectObject item.hdc, m_fontManager.getDefaultFont
          
          Dim textRect As RECT
          
180       textRect = item.rcItem
190       textRect.right = textRect.left + COLUMN_NAME
          
200       swiftTextOut item.hdc, textRect.left, textRect.top, ETO_CLIPPED, VarPtr(textRect), channel.name

210       textRect.left = textRect.left + COLUMN_NAME + COLUMN_SPACING
220       textRect.right = textRect.left + COLUMN_USERS
          
          Dim userCount As String
          
230       userCount = CStr(channel.userCount)
          
240       swiftTextOut item.hdc, textRect.left, textRect.top, ETO_CLIPPED, VarPtr(textRect), userCount
          
250       textRect.left = textRect.right + COLUMN_SPACING
          
          Dim totalWidth As Long
          
260       totalWidth = drawTopic(item.hdc, textRect, selected, channel.topic)
          
270       If totalWidth > m_maxTopicWidth Then
280           m_maxTopicWidth = totalWidth
290           SendMessage m_listbox, LB_SETHORIZONTALEXTENT, m_maxTopicWidth, ByVal 0
300       End If
          
310       drawItem = 1
End Function

Private Function drawTopic(hdc As Long, textRect As RECT, selected As Boolean, topic As String) As Long
          Dim blocks As New cArrayList
          
10        parseIrcFormatting topic, blocks
          
          Dim block As ITextRenderBlock
          Dim blockText As CBlockText
          Dim blockBold As CBlockBold
          Dim blockItalic As CBlockItalic
          Dim blockUnderline As CBlockUnderline
          Dim blockReverse As CBlockReverse
          Dim blockForeColour As CBlockForeColour
          Dim blockbackColour As CBlockBackColour
          
          Dim textSize As SIZE
          Dim width As Long
          
          Dim bold As Boolean
          Dim italic As Boolean
          Dim underline As Boolean
          Dim reverse As Boolean
          
          Dim drawingData As CDrawingData
          
          Dim foreColour As Long
          Dim backColour As Long
          
20        backColour = g_channelListBack
30        foreColour = g_channelListFore
          
          Dim count As Long
          
40        For count = 1 To blocks.count
50            Set block = blocks.item(count)
          
60            If TypeOf block Is CBlockText Then
70                Set blockText = block
                  
80                swiftTextOut hdc, textRect.left, textRect.top, 0, ByVal 0, blockText.text
                  
90                swiftGetTextExtentPoint32 hdc, blockText.text, textSize
100               textRect.left = textRect.left + textSize.cx
110           ElseIf TypeOf block Is CBlockBold Then
120               Set blockBold = block
130               bold = blockBold.bold
140               SelectObject hdc, m_fontManager.getFont(bold, italic, underline)
150           ElseIf TypeOf block Is CBlockItalic Then
160               Set blockItalic = block
170               italic = blockItalic.italic
180               SelectObject hdc, m_fontManager.getFont(bold, italic, underline)
190           ElseIf TypeOf block Is CBlockUnderline Then
200               Set blockUnderline = block
210               underline = blockUnderline.underline
220               SelectObject hdc, m_fontManager.getFont(bold, italic, underline)
230           ElseIf TypeOf block Is CBlockReverse Then
240               If Not selected Then
250                   Set blockReverse = block
                      
260                   reverse = blockReverse.reverse
                      
270                   If reverse Then
280                       SetBkColor hdc, getPaletteEntry(g_channelListFore)
290                       SetTextColor hdc, getPaletteEntry(g_channelListBack)
300                   Else
310                       SetBkColor hdc, getPaletteEntry(backColour)
320                       SetTextColor hdc, getPaletteEntry(foreColour)
330                   End If
340               End If
350           ElseIf TypeOf block Is CBlockForeColour Then
360               If Not selected Then
370                   Set blockForeColour = block
                      
380                   If blockForeColour.hasForeColour Then
390                       foreColour = blockForeColour.foreColour
                          
400                       If Not reverse Then
410                           SetTextColor hdc, getPaletteEntry(foreColour)
420                       End If
430                   Else
440                       foreColour = g_channelListFore
                      
450                       If Not reverse Then
460                           SetTextColor hdc, getPaletteEntry(foreColour)
470                       End If
480                   End If
490               End If
500           ElseIf TypeOf block Is CBlockBackColour Then
510               If Not selected Then
520                   Set blockbackColour = block
                      
530                   If blockbackColour.hasBackColour Then
540                       backColour = blockbackColour.backColour
                          
550                       If Not reverse Then
560                           SetBkColor hdc, getPaletteEntry(backColour)
570                       End If
580                   Else
590                       backColour = g_channelListBack
                      
600                       If Not reverse Then
610                           SetBkColor hdc, getPaletteEntry(backColour)
620                       End If
630                   End If
640               End If
650           ElseIf TypeOf block Is CBlockNormal Then
660               bold = False
670               italic = False
680               underline = False
690               reverse = False

700               foreColour = g_channelListFore
710               backColour = g_channelListBack
                  
720               SelectObject hdc, m_fontManager.getDefaultFont
                  
730               If Not selected Then
740                   SetBkColor hdc, getPaletteEntry(backColour)
750                   SetTextColor hdc, getPaletteEntry(foreColour)
760               End If
770           End If
780       Next count
          
790       drawTopic = textRect.left
End Function

Private Function isFormatCode(wChar As Integer) As Boolean
10        isFormatCode = True

20        Select Case wChar
              Case 2
30            Case 3
40            Case 4
50            Case 15
60            Case 22
70            Case 31
80            Case Else
90                isFormatCode = False
100       End Select
End Function

Private Sub parseIrcFormatting(text As String, ByRef blocks As cArrayList)
          Dim count As Integer
          Dim wChar As Integer
          Dim last As Integer
          Dim Length As Integer
          
          Dim blockText As CBlockText
          Dim blockBold As CBlockBold
          Dim blockItalic As CBlockItalic
          Dim blockUnderline As CBlockUnderline
          Dim blockReverse As CBlockReverse
          Dim blockForeColour As CBlockForeColour
          Dim blockbackColour As CBlockBackColour
          
          Dim toggleBold As Boolean
          Dim toggleItalic As Boolean
          Dim toggleUnderline As Boolean
          Dim toggleReverse As Boolean
          
10        last = 1
          
20        For count = 1 To Len(text)
30            wChar = AscW(Mid$(text, count, 1))
              
40            If isFormatCode(wChar) Then
50                Length = count - last
                  
60                If Length > 0 Then
70                    Set blockText = New CBlockText
                      
80                    blockText.text = Mid$(text, last, Length)
90                    blocks.Add blockText
100               End If
                  
110               Select Case wChar
                      Case 2
120                       toggleBold = Not toggleBold
130                       Set blockBold = New CBlockBold
140                       blockBold.bold = toggleBold
150                       blocks.Add blockBold
160                   Case 4
170                       toggleItalic = Not toggleItalic
180                       Set blockItalic = New CBlockItalic
190                       blockItalic.italic = toggleItalic
200                       blocks.Add blockItalic
210                   Case 15
220                       toggleBold = False
230                       toggleItalic = False
240                       toggleUnderline = False
250                       toggleReverse = False
                          
260                       blocks.Add New CBlockNormal
270                   Case 22
280                       toggleReverse = Not toggleReverse
290                       Set blockReverse = New CBlockReverse
300                       blockReverse.reverse = toggleReverse
310                       blocks.Add blockReverse
320                   Case 31
330                       toggleUnderline = Not toggleUnderline
340                       Set blockUnderline = New CBlockUnderline
350                       blockUnderline.underline = toggleUnderline
360                       blocks.Add blockUnderline
370                   Case 3
380                       If count = Len(text) Then
390                           blocks.Add New CBlockForeColour
400                           blocks.Add New CBlockBackColour
410                           Exit Sub
420                       End If
                          
                          Dim foreColour As Byte
                          Dim backColour As Byte
                          
430                       count = parseColourCode(text, count + 1, foreColour, backColour) - 1
                          
440                       If foreColour <> 255 Then
450                           Set blockForeColour = New CBlockForeColour
                              
460                           blockForeColour.hasForeColour = True
470                           blockForeColour.foreColour = foreColour
480                           blocks.Add blockForeColour
                              
490                       Else
500                           blocks.Add New CBlockForeColour
                              
510                           If backColour = 255 Then
520                               blocks.Add New CBlockBackColour
530                           End If
540                       End If
                          
550                       If backColour <> 255 Then
560                           Set blockbackColour = New CBlockBackColour
                              
570                           blockbackColour.hasBackColour = True
580                           blockbackColour.backColour = backColour
590                           blocks.Add blockbackColour
600                       End If
610               End Select
                  
620               last = count + 1
630           End If
640       Next count
          
650       If count - last > 0 Then
660           Set blockText = New CBlockText
              
670           blockText.text = Mid$(text, last)
680           blocks.Add blockText
690       End If
End Sub

Private Function parseColourCode(text As String, start As Integer, ByRef fore As Byte, ByRef back _
    As Byte) As Integer
          
          Dim colourCount As Integer
          Dim digits As Byte
          Dim currentColour As Byte
          Dim hasColour As Boolean
          
          Dim wChar As Integer
          
10        fore = 255
20        back = 255
          
30        For colourCount = start To Len(text)
40            wChar = AscW(Mid$(text, colourCount, 1))
              
50            If wChar > 47 And wChar < 58 Then
60                If digits = 0 Then
70                    hasColour = True
80                    currentColour = wChar - 48
90                    digits = 1
100                   start = start + 1
110               ElseIf digits = 1 Then
120                   currentColour = (currentColour * 10) + (wChar - 48)
130                   digits = 2
140                   start = start + 1
150               Else
160                   Exit For
170               End If
180           ElseIf wChar = AscW(",") Then
190               If Not hasColour Then
200                   Exit For
210               End If
                  
220               fore = currentColour
                  
230               If fore > 15 Then
240                   fore = fore Mod 16
250               End If
                  
260               hasColour = False
270               digits = 0
280               start = start + 1
290           Else
300               Exit For
310           End If
320       Next colourCount

330       If fore <> 255 Then
340           If hasColour Then
350               back = currentColour
                  
360               If back > 15 Then
370                   back = back Mod 16
380               End If
390           End If
400       ElseIf hasColour Then
410           fore = currentColour
              
420           If fore > 15 Then
430               fore = fore Mod 16
440           End If
450       End If
          
460       parseColourCode = start
End Function

Private Property Get ITabWindow_getTab() As CTab
10        Set ITabWindow_getTab = m_switchbarTab
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub mnuJoin_Click()
          Dim index As Long
          
10        index = SendMessage(m_listbox, LB_GETCURSEL, 0, ByVal 0)
20        RaiseEvent joinChannel(m_channels.item(index + 1).name)
End Sub

Private Sub mnuSortName_Click()
10        m_sortBy = estName
20        mnuSortName.Checked = True
30        mnuSortUsers.Checked = False
40        mnuSortTopic.Checked = False
50        sortList
End Sub

Private Sub mnuSortTopic_Click()
10        m_sortBy = estTopic
20        mnuSortTopic.Checked = True
30        mnuSortName.Checked = False
40        mnuSortUsers.Checked = False
50        sortList
End Sub

Private Sub mnuSortUsers_Click()
10        m_sortBy = estUsers
20        mnuSortUsers.Checked = True
30        mnuSortName.Checked = False
40        mnuSortTopic.Checked = False
50        sortList
End Sub

Private Sub UserControl_Initialize()
10        initMessages

          Dim IPAO As IOleInPlaceActiveObject

20        With m_IPAOHookStruct
30           Set IPAO = Me
40           CopyMemory .IPAOReal, IPAO, 4
50           CopyMemory .TBEx, Me, 4
60           .lpVTable = IPAOVTableChannelList
70           .ThisPointer = VarPtr(m_IPAOHookStruct)
80        End With

90        m_listbox = CreateWindowEx(0, "LISTBOX", 0&, WS_CHILD Or _
              WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL Or LBS_HASSTRINGS _
              Or LBS_OWNERDRAWFIXED Or LBS_NOINTEGRALHEIGHT, _
              0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
              UserControl.hwnd, 0&, App.hInstance, 0&)
          
100       AttachMessage Me, m_listbox, WM_MOUSEWHEEL
110       AttachMessage Me, m_listbox, WM_LBUTTONDBLCLK
120       AttachMessage Me, m_listbox, WM_RBUTTONDOWN
          
130       m_labelManager.addLabel "Channel name", ltSubHeading, 5, 0
140       m_labelManager.addLabel "User count", ltSubHeading, 130, 0
150       m_labelManager.addLabel "Topic", ltSubHeading, 150, 0
          
160       m_sortBy = estUsers
170       mnuSortUsers.Checked = True
End Sub

Friend Sub clear()
10        SendMessage m_listbox, LB_RESETCONTENT, 0, ByVal 0
20        Set m_channels = New cArrayList
End Sub

Friend Sub insertChannel(name As String, userCount As Long, topic As String)
          Dim tempName As String
          Dim tempTopic As String
          
10        If m_channels.count < 1 Then
20            addItem 1, name, userCount, topic
30            Exit Sub
40        End If
          
50        tempName = LCase$(name)
60        tempTopic = LCase$(topic)
          
          Dim pos As Long
          
70        If m_sortBy = estUsers Then
80            pos = findInsertPositionUsers(userCount, tempName)
90        ElseIf m_sortBy = estName Then
100           pos = findInsertPositionName(tempName)
110       ElseIf m_sortBy = estTopic Then
120           pos = findInsertPositionTopic(tempTopic)
130       End If
          
140       addItem pos, name, userCount, topic
End Sub

Private Sub sortList()
10        If m_channels.count = 0 Then
20            Exit Sub
30        End If

          Dim tempList As New cArrayList
          
40        Set tempList = m_channels
50        Set m_channels = New cArrayList
          
          Dim count As Long
          Dim index As Long
          
60        m_channels.Add tempList.item(1)
          
70        For count = 2 To tempList.count
80            If m_sortBy = estUsers Then
90                index = findInsertPositionUsers(tempList.item(count).userCount, LCase$(tempList.item(count).name))
100           ElseIf m_sortBy = estName Then
110               index = findInsertPositionName(LCase$(tempList.item(count).name))
120           ElseIf m_sortBy = estTopic Then
130               index = findInsertPositionTopic(LCase$(tempList.item(count).topic))
140           End If
              
150           m_channels.Add tempList.item(count), index
160       Next count
          
170       SendMessage m_listbox, LB_RESETCONTENT, 0, ByVal 0
          
180       For count = 1 To m_channels.count
190           SendMessage m_listbox, LB_ADDSTRING, 0, m_channels.item(count).name
200       Next count
End Sub

Private Function findInsertPositionName(name As String) As Long
          Dim min As Long
          Dim max As Long
          Dim pivotIndex As Long
          Dim pivot As String
          Dim result As Long
          
10        min = 1
20        max = m_channels.count

30        Do While (max - min) > 0
40            pivotIndex = (((max - min)) / 2) + min
50            pivot = m_channels.item(pivotIndex).name
              
60            result = StrComp(name, LCase$(m_channels.item(pivotIndex).name), vbBinaryCompare)
                      
70            If result = 1 Then
80                min = pivotIndex + 1
90            ElseIf result - 1 Then
100               max = pivotIndex - 1
110           Else
120               min = pivotIndex + 1
130           End If
140       Loop
          
          Dim pos As Long
          
150       result = StrComp(name, LCase$(m_channels.item(max).name))
          
160       If result = 1 Then
170           pos = max + 1
180       ElseIf result = -1 Then
190           pos = max
200       Else
210           pos = max + 1
220       End If
          
230       findInsertPositionName = pos
End Function

Private Function findInsertPositionUsers(users As Long, name As String)
          Dim min As Long
          Dim max As Long
          Dim pivotIndex As Long
          Dim pivot As Long
          Dim result As Long
          
10        min = 1
20        max = m_channels.count
          
30        Do While (max - min) > 0
40            pivotIndex = (((max - min)) / 2) + min
50            pivot = m_channels.item(pivotIndex).userCount
              
60            If users > pivot Then
70                max = pivotIndex - 1
80            ElseIf users < pivot Then
90                min = pivotIndex + 1
100           Else
110               result = StrComp(name, LCase$(m_channels.item(pivotIndex).name))
                      
120               If result = 1 Then
130                   min = pivotIndex + 1
140               ElseIf result - 1 Then
150                   max = pivotIndex - 1
160               Else
170                   min = pivotIndex + 1
180               End If
190           End If
200       Loop
          
          Dim pos As Long
          
210       pivot = m_channels.item(max).userCount
          
220       If users > pivot Then
230           pos = max
240       ElseIf users < pivot Then
250           pos = max + 1
260       Else
270           result = StrComp(name, LCase$(m_channels.item(max).name))
              
280           If result = 1 Then
290               pos = max + 1
300           ElseIf result = -1 Then
310               pos = max
320           Else
330               pos = max + 1
340           End If
350       End If
          
360       findInsertPositionUsers = pos
End Function

Private Function findInsertPositionTopic(topic As String) As Long
          Dim min As Long
          Dim max As Long
          Dim pivotIndex As Long
          Dim pivot As String
          Dim result As Long
          
10        min = 1
20        max = m_channels.count

30        Do While (max - min) > 0
40            pivotIndex = (((max - min)) / 2) + min
50            pivot = m_channels.item(pivotIndex).topic
              
60            result = StrComp(topic, LCase$(m_channels.item(pivotIndex).topic), vbBinaryCompare)
                      
70            If result = 1 Then
80                min = pivotIndex + 1
90            ElseIf result - 1 Then
100               max = pivotIndex - 1
110           Else
120               min = pivotIndex + 1
130           End If
140       Loop
          
          Dim pos As Long
          
150       result = StrComp(topic, LCase$(m_channels.item(max).topic))
          
160       If result = 1 Then
170           pos = max + 1
180       ElseIf result = -1 Then
190           pos = max
200       Else
210           pos = max + 1
220       End If
          
230       findInsertPositionTopic = pos
End Function

Private Sub addItem(index As Long, name As String, users As Long, topic As String)
          Dim channel As New CChannelListItem
10        channel.name = name
20        channel.userCount = users
30        channel.topic = topic
          
40        m_channels.Add channel, index
50        SendMessage m_listbox, LB_INSERTSTRING, index - 1, ByVal Mid$(name, 2)
60        drawChannelCount
End Sub

Private Sub UserControl_Paint()
10        SetBkColor UserControl.hdc, getPaletteEntry(g_channelListBack)
20        SetTextColor UserControl.hdc, getPaletteEntry(g_channelListFore)
          
30        UserControl.BackColor = getPaletteEntry(g_channelListBack)
40        SelectObject UserControl.hdc, m_fontManager.getDefaultFont
          
          Dim fillBrush As Long
          
50        fillBrush = CreateSolidBrush(getPaletteEntry(g_channelListBack))
60        FillRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, m_itemHeight), fillBrush
70        DeleteObject fillBrush
          
          Dim textSize As SIZE
          
80        swiftGetTextExtentPoint32 UserControl.hdc, "Channel name", textSize
90        COLUMN_NAME = textSize.cx
100       swiftGetTextExtentPoint32 UserControl.hdc, "Channel name", textSize
110       COLUMN_USERS = textSize.cx

120       swiftTextOut UserControl.hdc, 0, 0, 0, ByVal 0, "Channel name"
130       swiftTextOut UserControl.hdc, COLUMN_NAME + COLUMN_SPACING, 0, 0, ByVal 0, "User count"
140       swiftTextOut UserControl.hdc, COLUMN_NAME + COLUMN_USERS + (COLUMN_SPACING * 2), 0, 0, _
              ByVal 0, "Channel Topic"
          
150       drawChannelCount
End Sub

Private Sub drawChannelCount()
          Dim text As String
          Dim textSize As SIZE
          
10        text = m_channels.count & " channels"
20        swiftGetTextExtentPoint32 UserControl.hdc, text, textSize
          
30        SetBkMode UserControl.hdc, OPAQUE
          
40        swiftTextOut UserControl.hdc, UserControl.ScaleWidth - (textSize.cx + 5), 0, ETO_OPAQUE, ByVal 0, text
End Sub

Private Sub UserControl_Resize()
10        MoveWindow m_listbox, 0, m_itemHeight, UserControl.ScaleWidth, UserControl.ScaleHeight - m_itemHeight, 1
End Sub

Private Sub initMessages()
10        AttachMessage Me, UserControl.hwnd, WM_SETFOCUS
20        AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
30        AttachMessage Me, UserControl.hwnd, WM_MEASUREITEM
40        AttachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
50        AttachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
End Sub

Private Sub deInitMessages()
10        DetachMessage Me, UserControl.hwnd, WM_SETFOCUS
20        DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
30        DetachMessage Me, UserControl.hwnd, WM_MEASUREITEM
40        DetachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
50        DetachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
60        DetachMessage Me, m_listbox, WM_MOUSEWHEEL
70        DetachMessage Me, m_listbox, WM_LBUTTONDBLCLK
80        DetachMessage Me, m_listbox, WM_RBUTTONDOWN
End Sub

Private Sub UserControl_Terminate()
10        With m_IPAOHookStruct
20          CopyMemory .IPAOReal, 0&, 4
30          CopyMemory .TBEx, 0&, 4
40        End With
          
50        deInitMessages
End Sub
