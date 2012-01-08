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

Private m_fontmanager As CFontManager
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
    Set session = m_session
End Property

Public Property Let session(newValue As CSession)
    Set m_session = newValue
End Property

Public Property Get switchbartab() As CTab
    Set switchbartab = m_switchbarTab
End Property

Public Property Let switchbartab(newValue As CTab)
    Set m_switchbarTab = newValue
End Property

Private Sub IColourUser_coloursUpdated()

End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontmanager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
    fontsUpdated
End Sub

Private Sub fontsUpdated()
    m_itemHeight = m_fontmanager.fontHeight
    UserControl_Resize
    
    If Not m_listbox = 0 Then
        SendMessage m_listbox, LB_SETITEMHEIGHT, 0, ByVal m_itemHeight
        
        If m_channels.count > 0 Then
            Dim TopIndex As Long
            
            TopIndex = SendMessage(m_listbox, LB_GETTOPINDEX, 0, ByVal 0)
            
            If TopIndex > (m_channels.count - 1) - (UserControl.ScaleHeight / m_itemHeight) Then
                'If the font change has left space at the bottom of
                'the listbox, update the scroll index appropriately.
                SendMessage m_listbox, LB_SETTOPINDEX, m_channels.count - 1, ByVal 0
            End If
        End If
    End If
End Sub

Public Function TranslateAccelerator(lpMsg As VBOleGuids.Msg) As Long
    TranslateAccelerator = S_FALSE
    ' Here you can modify the response to the key down
    ' accelerator command using the values in lpMsg.  This
    ' can be used to capture Tabs, Returns, Arrows etc.
    ' Just process the message as required and return S_OK.
    Dim key As Long
    
    key = (lpMsg.wParam And &HFFFF&)
    
    If key = vbKeyLeft Or key = vbKeyRight Or key = vbKeyUp Or key = vbKeyDown Then
        If lpMsg.message = WM_KEYDOWN Or lpMsg.message = WM_KEYUP Then
            SendMessage m_listbox, lpMsg.message, key, 0
            TranslateAccelerator = S_OK
        End If
    End If
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_MOUSEWHEEL
            ISubclass_MsgResponse = emrConsume
        Case WM_LBUTTONDBLCLK
            ISubclass_MsgResponse = emrPostProcess
        Case Else
            ISubclass_MsgResponse = emrPreprocess
    End Select
End Property

Private Property Get ShiftState() As Integer
    If GetAsyncKeyState(vbKeyShift) <> 0 Then
        ShiftState = vbShiftMask
    End If

    If GetAsyncKeyState(vbKeyControl) <> 0 Then
        ShiftState = ShiftState Or vbCtrlMask
    End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim coordY As Long
    
    Select Case iMsg
        Case WM_DRAWITEM
            Dim item As DRAWITEMSTRUCT
            
            CopyMemory item, ByVal lParam, Len(item)
            ISubclass_WindowProc = drawItem(item)
        Case WM_MEASUREITEM
            Dim measureItem As MEASUREITEMSTRUCT
            
            CopyMemory measureItem, ByVal lParam, Len(measureItem)
            measureItem.itemHeight = m_itemHeight
            CopyMemory ByVal lParam, measureItem, Len(measureItem)
        
            ISubclass_WindowProc = True
        Case WM_SETFOCUS
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
            Set pOleObject = Me
            Set pOleInPlaceSite = pOleObject.GetClientSite
            pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), _
                VarPtr(ClipRect), VarPtr(FrameInfo)
            CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
            pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
            If Not pOleInPlaceUIWindow Is Nothing Then
               pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
            End If
            ' Clear up the inbetween implementation:
            CopyMemory pOleInPlaceActiveObject, 0&, 4
            ' --------------------------------------------------------------------------
        Case WM_CTLCOLORLISTBOX
            If m_fillbrush <> 0 Then
                DeleteObject m_fillbrush
            End If
            
            m_fillbrush = CreateSolidBrush(getPaletteEntry(g_channelListBack))
            ISubclass_WindowProc = m_fillbrush
        Case WM_MOUSEWHEEL
            If Not m_listbox = 0 Then
                Dim accumDelta As Integer
    
                accumDelta = HiWord(wParam)
        
                Do While accumDelta >= 40
                    SendMessage m_listbox, WM_VSCROLL, SB_LINEUP, 0
                    accumDelta = accumDelta - 40
                Loop
        
                Do While accumDelta <= -40
                    SendMessage m_listbox, WM_VSCROLL, SB_LINEDOWN, 0
                    accumDelta = accumDelta + 40
                Loop
            End If
        Case WM_LBUTTONDBLCLK
            Dim index As Long
            
            index = SendMessage(m_listbox, LB_GETCURSEL, 0, ByVal 0)
            
            If index <> LB_ERR Then
                RaiseEvent joinChannel(m_channels.item(index + 1).name)
            End If
        Case WM_RBUTTONDOWN
            coordY = HiWord(lParam)
            
            Dim TopIndex As Long
            Dim itemIndex As Long
            
            TopIndex = SendMessage(m_listbox, LB_GETTOPINDEX, 0, ByVal 0)
            itemIndex = TopIndex + Fix(coordY / m_itemHeight)
            
            SendMessage m_listbox, LB_SETCURSEL, itemIndex, ByVal 0
            PopupMenu mnuChannelListContext
    End Select
End Function

Private Function drawItem(item As DRAWITEMSTRUCT) As Long
    If item.itemID = -1 Then
        Exit Function
    End If
    
    Dim channel As CChannelListItem
    Dim selected As Boolean

    Set channel = m_channels.item(item.itemID + 1)
    
    selected = item.itemState And ODS_SELECTED
    
    If selected Then
        SetBkColor item.hdc, GetSysColor(COLOR_HIGHLIGHT)
        SetTextColor item.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
        FillRect item.hdc, item.rcItem, GetSysColorBrush(COLOR_HIGHLIGHT)
    Else
        SetBkColor item.hdc, getPaletteEntry(g_channelListBack)
        SetTextColor item.hdc, getPaletteEntry(g_channelListFore)
        
        Dim fillBrush As Long
        fillBrush = CreateSolidBrush(getPaletteEntry(g_channelListBack))
        FillRect item.hdc, item.rcItem, fillBrush
        DeleteObject fillBrush
    End If
    
    SelectObject item.hdc, m_fontmanager.getDefaultFont
    
    Dim textRect As RECT
    
    textRect = item.rcItem
    textRect.right = textRect.left + COLUMN_NAME
    
    swiftTextOut item.hdc, textRect.left, textRect.top, ETO_CLIPPED, VarPtr(textRect), channel.name

    textRect.left = textRect.left + COLUMN_NAME + COLUMN_SPACING
    textRect.right = textRect.left + COLUMN_USERS
    
    Dim userCount As String
    
    userCount = CStr(channel.userCount)
    
    swiftTextOut item.hdc, textRect.left, textRect.top, ETO_CLIPPED, VarPtr(textRect), userCount
    
    textRect.left = textRect.right + COLUMN_SPACING
    
    Dim totalWidth As Long
    
    totalWidth = drawTopic(item.hdc, textRect, selected, channel.topic)
    
    If totalWidth > m_maxTopicWidth Then
        m_maxTopicWidth = totalWidth
        SendMessage m_listbox, LB_SETHORIZONTALEXTENT, m_maxTopicWidth, ByVal 0
    End If
    
    drawItem = 1
End Function

Private Function drawTopic(hdc As Long, textRect As RECT, selected As Boolean, topic As String) As Long
    Dim blocks As New cArrayList
    
    parseIrcFormatting topic, blocks
    
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
    
    backColour = g_channelListBack
    foreColour = g_channelListFore
    
    Dim count As Long
    
    For count = 1 To blocks.count
        Set block = blocks.item(count)
    
        If TypeOf block Is CBlockText Then
            Set blockText = block
            
            swiftTextOut hdc, textRect.left, textRect.top, 0, ByVal 0, blockText.text
            
            swiftGetTextExtentPoint32 hdc, blockText.text, textSize
            textRect.left = textRect.left + textSize.cx
        ElseIf TypeOf block Is CBlockBold Then
            Set blockBold = block
            bold = blockBold.bold
            SelectObject hdc, m_fontmanager.getFont(bold, italic, underline)
        ElseIf TypeOf block Is CBlockItalic Then
            Set blockItalic = block
            italic = blockItalic.italic
            SelectObject hdc, m_fontmanager.getFont(bold, italic, underline)
        ElseIf TypeOf block Is CBlockUnderline Then
            Set blockUnderline = block
            underline = blockUnderline.underline
            SelectObject hdc, m_fontmanager.getFont(bold, italic, underline)
        ElseIf TypeOf block Is CBlockReverse Then
            If Not selected Then
                Set blockReverse = block
                
                reverse = blockReverse.reverse
                
                If reverse Then
                    SetBkColor hdc, getPaletteEntry(g_channelListFore)
                    SetTextColor hdc, getPaletteEntry(g_channelListBack)
                Else
                    SetBkColor hdc, getPaletteEntry(backColour)
                    SetTextColor hdc, getPaletteEntry(foreColour)
                End If
            End If
        ElseIf TypeOf block Is CBlockForeColour Then
            If Not selected Then
                Set blockForeColour = block
                
                If blockForeColour.hasForeColour Then
                    foreColour = blockForeColour.foreColour
                    
                    If Not reverse Then
                        SetTextColor hdc, getPaletteEntry(foreColour)
                    End If
                Else
                    foreColour = g_channelListFore
                
                    If Not reverse Then
                        SetTextColor hdc, getPaletteEntry(foreColour)
                    End If
                End If
            End If
        ElseIf TypeOf block Is CBlockBackColour Then
            If Not selected Then
                Set blockbackColour = block
                
                If blockbackColour.hasBackColour Then
                    backColour = blockbackColour.backColour
                    
                    If Not reverse Then
                        SetBkColor hdc, getPaletteEntry(backColour)
                    End If
                Else
                    backColour = g_channelListBack
                
                    If Not reverse Then
                        SetBkColor hdc, getPaletteEntry(backColour)
                    End If
                End If
            End If
        ElseIf TypeOf block Is CBlockNormal Then
            bold = False
            italic = False
            underline = False
            reverse = False

            foreColour = g_channelListFore
            backColour = g_channelListBack
            
            SelectObject hdc, m_fontmanager.getDefaultFont
            
            If Not selected Then
                SetBkColor hdc, getPaletteEntry(backColour)
                SetTextColor hdc, getPaletteEntry(foreColour)
            End If
        End If
    Next count
    
    drawTopic = textRect.left
End Function

Private Function isFormatCode(wChar As Integer) As Boolean
    isFormatCode = True

    Select Case wChar
        Case 2
        Case 3
        Case 4
        Case 15
        Case 22
        Case 31
        Case Else
            isFormatCode = False
    End Select
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
    
    last = 1
    
    For count = 1 To Len(text)
        wChar = AscW(Mid$(text, count, 1))
        
        If isFormatCode(wChar) Then
            Length = count - last
            
            If Length > 0 Then
                Set blockText = New CBlockText
                
                blockText.text = Mid$(text, last, Length)
                blocks.Add blockText
            End If
            
            Select Case wChar
                Case 2
                    toggleBold = Not toggleBold
                    Set blockBold = New CBlockBold
                    blockBold.bold = toggleBold
                    blocks.Add blockBold
                Case 4
                    toggleItalic = Not toggleItalic
                    Set blockItalic = New CBlockItalic
                    blockItalic.italic = toggleItalic
                    blocks.Add blockItalic
                Case 15
                    toggleBold = False
                    toggleItalic = False
                    toggleUnderline = False
                    toggleReverse = False
                    
                    blocks.Add New CBlockNormal
                Case 22
                    toggleReverse = Not toggleReverse
                    Set blockReverse = New CBlockReverse
                    blockReverse.reverse = toggleReverse
                    blocks.Add blockReverse
                Case 31
                    toggleUnderline = Not toggleUnderline
                    Set blockUnderline = New CBlockUnderline
                    blockUnderline.underline = toggleUnderline
                    blocks.Add blockUnderline
                Case 3
                    If count = Len(text) Then
                        blocks.Add New CBlockForeColour
                        blocks.Add New CBlockBackColour
                        Exit Sub
                    End If
                    
                    Dim foreColour As Byte
                    Dim backColour As Byte
                    
                    count = parseColourCode(text, count + 1, foreColour, backColour) - 1
                    
                    If foreColour <> 255 Then
                        Set blockForeColour = New CBlockForeColour
                        
                        blockForeColour.hasForeColour = True
                        blockForeColour.foreColour = foreColour
                        blocks.Add blockForeColour
                        
                    Else
                        blocks.Add New CBlockForeColour
                        
                        If backColour = 255 Then
                            blocks.Add New CBlockBackColour
                        End If
                    End If
                    
                    If backColour <> 255 Then
                        Set blockbackColour = New CBlockBackColour
                        
                        blockbackColour.hasBackColour = True
                        blockbackColour.backColour = backColour
                        blocks.Add blockbackColour
                    End If
            End Select
            
            last = count + 1
        End If
    Next count
    
    If count - last > 0 Then
        Set blockText = New CBlockText
        
        blockText.text = Mid$(text, last)
        blocks.Add blockText
    End If
End Sub

Private Function parseColourCode(text As String, start As Integer, ByRef fore As Byte, ByRef back _
    As Byte) As Integer
    
    Dim colourCount As Integer
    Dim digits As Byte
    Dim currentColour As Byte
    Dim hasColour As Boolean
    
    Dim wChar As Integer
    
    fore = 255
    back = 255
    
    For colourCount = start To Len(text)
        wChar = AscW(Mid$(text, colourCount, 1))
        
        If wChar > 47 And wChar < 58 Then
            If digits = 0 Then
                hasColour = True
                currentColour = wChar - 48
                digits = 1
                start = start + 1
            ElseIf digits = 1 Then
                currentColour = (currentColour * 10) + (wChar - 48)
                digits = 2
                start = start + 1
            Else
                Exit For
            End If
        ElseIf wChar = AscW(",") Then
            If Not hasColour Then
                Exit For
            End If
            
            fore = currentColour
            
            If fore > 15 Then
                fore = fore Mod 16
            End If
            
            hasColour = False
            digits = 0
            start = start + 1
        Else
            Exit For
        End If
    Next colourCount

    If fore <> 255 Then
        If hasColour Then
            back = currentColour
            
            If back > 15 Then
                back = back Mod 16
            End If
        End If
    ElseIf hasColour Then
        fore = currentColour
        
        If fore > 15 Then
            fore = fore Mod 16
        End If
    End If
    
    parseColourCode = start
End Function

Private Property Get ITabWindow_getTab() As CTab
    Set ITabWindow_getTab = m_switchbarTab
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub mnuJoin_Click()
    Dim index As Long
    
    index = SendMessage(m_listbox, LB_GETCURSEL, 0, ByVal 0)
    RaiseEvent joinChannel(m_channels.item(index + 1).name)
End Sub

Private Sub mnuSortName_Click()
    m_sortBy = estName
    mnuSortName.Checked = True
    mnuSortUsers.Checked = False
    mnuSortTopic.Checked = False
    sortList
End Sub

Private Sub mnuSortTopic_Click()
    m_sortBy = estTopic
    mnuSortTopic.Checked = True
    mnuSortName.Checked = False
    mnuSortUsers.Checked = False
    sortList
End Sub

Private Sub mnuSortUsers_Click()
    m_sortBy = estUsers
    mnuSortUsers.Checked = True
    mnuSortName.Checked = False
    mnuSortTopic.Checked = False
    sortList
End Sub

Private Sub UserControl_Initialize()
    initMessages

    Dim IPAO As IOleInPlaceActiveObject

    With m_IPAOHookStruct
       Set IPAO = Me
       CopyMemory .IPAOReal, IPAO, 4
       CopyMemory .TBEx, Me, 4
       .lpVTable = IPAOVTableChannelList
       .ThisPointer = VarPtr(m_IPAOHookStruct)
    End With

    m_listbox = CreateWindowEx(0, "LISTBOX", 0&, WS_CHILD Or _
        WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL Or LBS_HASSTRINGS _
        Or LBS_OWNERDRAWFIXED Or LBS_NOINTEGRALHEIGHT, _
        0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        UserControl.hwnd, 0&, App.hInstance, 0&)
    
    AttachMessage Me, m_listbox, WM_MOUSEWHEEL
    AttachMessage Me, m_listbox, WM_LBUTTONDBLCLK
    AttachMessage Me, m_listbox, WM_RBUTTONDOWN
    
    m_labelManager.addLabel "Channel name", ltSubHeading, 5, 0
    m_labelManager.addLabel "User count", ltSubHeading, 130, 0
    m_labelManager.addLabel "Topic", ltSubHeading, 150, 0
    
    m_sortBy = estUsers
    mnuSortUsers.Checked = True
End Sub

Friend Sub clear()
    SendMessage m_listbox, LB_RESETCONTENT, 0, ByVal 0
    Set m_channels = New cArrayList
End Sub

Friend Sub insertChannel(name As String, userCount As Long, topic As String)
    Dim tempName As String
    Dim tempTopic As String
    
    If m_channels.count < 1 Then
        addItem 1, name, userCount, topic
        Exit Sub
    End If
    
    tempName = LCase$(name)
    tempTopic = LCase$(topic)
    
    Dim pos As Long
    
    If m_sortBy = estUsers Then
        pos = findInsertPositionUsers(userCount, tempName)
    ElseIf m_sortBy = estName Then
        pos = findInsertPositionName(tempName)
    ElseIf m_sortBy = estTopic Then
        pos = findInsertPositionTopic(tempTopic)
    End If
    
    addItem pos, name, userCount, topic
End Sub

Private Sub sortList()
    If m_channels.count = 0 Then
        Exit Sub
    End If

    Dim tempList As New cArrayList
    
    Set tempList = m_channels
    Set m_channels = New cArrayList
    
    Dim count As Long
    Dim index As Long
    
    m_channels.Add tempList.item(1)
    
    For count = 2 To tempList.count
        If m_sortBy = estUsers Then
            index = findInsertPositionUsers(tempList.item(count).userCount, LCase$(tempList.item(count).name))
        ElseIf m_sortBy = estName Then
            index = findInsertPositionName(LCase$(tempList.item(count).name))
        ElseIf m_sortBy = estTopic Then
            index = findInsertPositionTopic(LCase$(tempList.item(count).topic))
        End If
        
        m_channels.Add tempList.item(count), index
    Next count
    
    SendMessage m_listbox, LB_RESETCONTENT, 0, ByVal 0
    
    For count = 1 To m_channels.count
        SendMessage m_listbox, LB_ADDSTRING, 0, m_channels.item(count).name
    Next count
End Sub

Private Function findInsertPositionName(name As String) As Long
    Dim min As Long
    Dim max As Long
    Dim pivotIndex As Long
    Dim pivot As String
    Dim result As Long
    
    min = 1
    max = m_channels.count

    Do While (max - min) > 0
        pivotIndex = (((max - min)) / 2) + min
        pivot = m_channels.item(pivotIndex).name
        
        result = StrComp(name, LCase$(m_channels.item(pivotIndex).name), vbBinaryCompare)
                
        If result = 1 Then
            min = pivotIndex + 1
        ElseIf result - 1 Then
            max = pivotIndex - 1
        Else
            min = pivotIndex + 1
        End If
    Loop
    
    Dim pos As Long
    
    result = StrComp(name, LCase$(m_channels.item(max).name))
    
    If result = 1 Then
        pos = max + 1
    ElseIf result = -1 Then
        pos = max
    Else
        pos = max + 1
    End If
    
    findInsertPositionName = pos
End Function

Private Function findInsertPositionUsers(users As Long, name As String)
    Dim min As Long
    Dim max As Long
    Dim pivotIndex As Long
    Dim pivot As Long
    Dim result As Long
    
    min = 1
    max = m_channels.count
    
    Do While (max - min) > 0
        pivotIndex = (((max - min)) / 2) + min
        pivot = m_channels.item(pivotIndex).userCount
        
        If users > pivot Then
            max = pivotIndex - 1
        ElseIf users < pivot Then
            min = pivotIndex + 1
        Else
            result = StrComp(name, LCase$(m_channels.item(pivotIndex).name))
                
            If result = 1 Then
                min = pivotIndex + 1
            ElseIf result - 1 Then
                max = pivotIndex - 1
            Else
                min = pivotIndex + 1
            End If
        End If
    Loop
    
    Dim pos As Long
    
    pivot = m_channels.item(max).userCount
    
    If users > pivot Then
        pos = max
    ElseIf users < pivot Then
        pos = max + 1
    Else
        result = StrComp(name, LCase$(m_channels.item(max).name))
        
        If result = 1 Then
            pos = max + 1
        ElseIf result = -1 Then
            pos = max
        Else
            pos = max + 1
        End If
    End If
    
    findInsertPositionUsers = pos
End Function

Private Function findInsertPositionTopic(topic As String) As Long
    Dim min As Long
    Dim max As Long
    Dim pivotIndex As Long
    Dim pivot As String
    Dim result As Long
    
    min = 1
    max = m_channels.count

    Do While (max - min) > 0
        pivotIndex = (((max - min)) / 2) + min
        pivot = m_channels.item(pivotIndex).topic
        
        result = StrComp(topic, LCase$(m_channels.item(pivotIndex).topic), vbBinaryCompare)
                
        If result = 1 Then
            min = pivotIndex + 1
        ElseIf result - 1 Then
            max = pivotIndex - 1
        Else
            min = pivotIndex + 1
        End If
    Loop
    
    Dim pos As Long
    
    result = StrComp(topic, LCase$(m_channels.item(max).topic))
    
    If result = 1 Then
        pos = max + 1
    ElseIf result = -1 Then
        pos = max
    Else
        pos = max + 1
    End If
    
    findInsertPositionTopic = pos
End Function

Private Sub addItem(index As Long, name As String, users As Long, topic As String)
    Dim channel As New CChannelListItem
    channel.name = name
    channel.userCount = users
    channel.topic = topic
    
    m_channels.Add channel, index
    SendMessage m_listbox, LB_INSERTSTRING, index - 1, ByVal Mid$(name, 2)
    drawChannelCount
End Sub

Private Sub UserControl_Paint()
    SetBkColor UserControl.hdc, getPaletteEntry(g_channelListBack)
    SetTextColor UserControl.hdc, getPaletteEntry(g_channelListFore)
    
    UserControl.BackColor = getPaletteEntry(g_channelListBack)
    SelectObject UserControl.hdc, m_fontmanager.getDefaultFont
    
    Dim fillBrush As Long
    
    fillBrush = CreateSolidBrush(getPaletteEntry(g_channelListBack))
    FillRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, m_itemHeight), fillBrush
    DeleteObject fillBrush
    
    Dim textSize As SIZE
    
    swiftGetTextExtentPoint32 UserControl.hdc, "Channel name", textSize
    COLUMN_NAME = textSize.cx
    swiftGetTextExtentPoint32 UserControl.hdc, "Channel name", textSize
    COLUMN_USERS = textSize.cx

    swiftTextOut UserControl.hdc, 0, 0, 0, ByVal 0, "Channel name"
    swiftTextOut UserControl.hdc, COLUMN_NAME + COLUMN_SPACING, 0, 0, ByVal 0, "User count"
    swiftTextOut UserControl.hdc, COLUMN_NAME + COLUMN_USERS + (COLUMN_SPACING * 2), 0, 0, _
        ByVal 0, "Channel Topic"
    
    drawChannelCount
End Sub

Private Sub drawChannelCount()
    Dim text As String
    Dim textSize As SIZE
    
    text = m_channels.count & " channels"
    swiftGetTextExtentPoint32 UserControl.hdc, text, textSize
    
    SetBkMode UserControl.hdc, OPAQUE
    
    swiftTextOut UserControl.hdc, UserControl.ScaleWidth - (textSize.cx + 5), 0, ETO_OPAQUE, ByVal 0, text
End Sub

Private Sub UserControl_Resize()
    MoveWindow m_listbox, 0, m_itemHeight, UserControl.ScaleWidth, UserControl.ScaleHeight - m_itemHeight, 1
End Sub

Private Sub initMessages()
    AttachMessage Me, UserControl.hwnd, WM_SETFOCUS
    AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
    AttachMessage Me, UserControl.hwnd, WM_MEASUREITEM
    AttachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
    AttachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
End Sub

Private Sub deInitMessages()
    DetachMessage Me, UserControl.hwnd, WM_SETFOCUS
    DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
    DetachMessage Me, UserControl.hwnd, WM_MEASUREITEM
    DetachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
    DetachMessage Me, UserControl.hwnd, WM_MOUSEWHEEL
    DetachMessage Me, m_listbox, WM_MOUSEWHEEL
    DetachMessage Me, m_listbox, WM_LBUTTONDBLCLK
    DetachMessage Me, m_listbox, WM_RBUTTONDOWN
End Sub

Private Sub UserControl_Terminate()
    With m_IPAOHookStruct
      CopyMemory .IPAOReal, 0&, 4
      CopyMemory .TBEx, 0&, 4
    End With
    
    deInitMessages
End Sub
