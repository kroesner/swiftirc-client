VERSION 5.00
Begin VB.UserControl ctlNickList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlNickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass
Implements IWindow
Implements IFontUser
Implements IColourUser

Public Event rightClicked(item As CNicklistItem)
Public Event doubleClicked(item As CNicklistItem)
Public Event changeWidth(newWidth As Long)

Private m_realWindow As VBControlExtender

Private m_listbox As ListBox
Private m_list As New cArrayList

Private m_fontManager As CFontManager
Private m_itemHeight As Long

Private m_resizing As Boolean

Private m_fillbrush As Long

Private m_session As CSession

Public Property Get item(index As Long) As CNicklistItem
    Set item = m_list.item(index)
End Property

Public Property Get itemCount() As Long
    itemCount = m_list.count
End Property

Public Property Get session() As CSession
    Set session = m_session
End Property

Public Property Let session(newValue As CSession)
    Set m_session = newValue
End Property

Private Sub IColourUser_coloursUpdated()
    refresh
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
    fontsUpdated
End Sub

Private Sub fontsUpdated()
    m_itemHeight = m_fontManager.fontHeight
    
    If Not m_listbox Is Nothing Then
        SendMessage m_listbox.hwnd, LB_SETITEMHEIGHT, 0, ByVal m_itemHeight
        
        If m_list.count > 0 Then
            Dim TopIndex As Long
            
            TopIndex = m_listbox.TopIndex
            
            If TopIndex > (m_list.count - 1) - (UserControl.ScaleHeight) / m_itemHeight Then
                'If the font change has left space at the bottom of
                'the listbox, update the scroll index appropriately.
                m_listbox.TopIndex = m_list.count - 1
            End If
        End If
        
        Me.refresh
    End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_DRAWITEM
            ISubclass_MsgResponse = emrConsume
        Case WM_MOUSEMOVE
            ISubclass_MsgResponse = emrPreprocess
        Case WM_LBUTTONDOWN
            ISubclass_MsgResponse = emrConsume
        Case Else
            ISubclass_MsgResponse = emrConsume
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    Dim count As Long
    Dim coordY As Long
    Dim coordX As Long
    Dim TopIndex As Long
    Dim itemIndex As Long

    Select Case iMsg
        Case WM_MOUSEMOVE
            coordX = LoWord(lParam)
        
            If coordX <= 0 Or coordX > 5 Then
                If m_resizing = False Then
                    m_listbox.MousePointer = vbArrow
                End If
            ElseIf coordX <= 5 Then
                m_listbox.MousePointer = vbSizeWE
            End If
            
            If m_resizing = True Then
                RaiseEvent changeWidth(UserControl.ScaleWidth - coordX)
            End If
        Case WM_LBUTTONDOWN
            coordX = LoWord(lParam)
        
            If coordX > 0 And coordX <= 5 Then
                m_resizing = True
                SetCapture m_listbox.hwnd
            Else
                CallOldWindowProc hwnd, iMsg, wParam, lParam
            End If
        Case WM_LBUTTONUP
            ReleaseCapture
            m_resizing = False
        Case WM_LBUTTONDBLCLK
            coordY = HiWord(lParam)
            
            TopIndex = m_listbox.TopIndex
            itemIndex = TopIndex + Fix(coordY / m_itemHeight)
            
            If itemIndex < m_listbox.ListCount Then
                RaiseEvent doubleClicked(m_list.item(itemIndex + 1))
            End If
        Case WM_RBUTTONDOWN
            coordY = HiWord(lParam)
            
            TopIndex = m_listbox.TopIndex
            itemIndex = TopIndex + Fix(coordY / m_itemHeight)
            
            If itemIndex < m_listbox.ListCount Then
                Dim selected As Long
                
                selected = SendMessage(m_listbox.hwnd, LB_GETSEL, itemIndex, ByVal 0&)
                
                'If the user right clicked on an unselected
                'item, we'll deselect any other items and select
                'only the clicked item.  Otherwise, we reselect
                'the clicked item to give it the focus caret.
                
                If selected = 0 Then
                    SendMessage m_listbox.hwnd, LB_SETSEL, 0, ByVal -1
                    SendMessage m_listbox.hwnd, LB_SETSEL, 1, ByVal itemIndex
                ElseIf selected > 0 Then
                    SendMessage m_listbox.hwnd, LB_SETSEL, 1, ByVal itemIndex
                End If
            End If
            
            Dim firstSelected As Long
            
            For count = 0 To m_listbox.ListCount - 1
                If m_listbox.selected(count) Then
                    firstSelected = count
                    Exit For
                End If
            Next count
            
            RaiseEvent rightClicked(m_list.item(firstSelected + 1))
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
        Case WM_CTLCOLORLISTBOX
            If m_fillbrush <> 0 Then
                DeleteObject m_fillbrush
            End If
        
            m_fillbrush = CreateSolidBrush(getPaletteEntry(g_nicklistBack))
            
            ISubclass_WindowProc = m_fillbrush
    End Select
End Function

Public Sub getSelectedItems(list As cArrayList)
    Dim count As Long
    
    For count = 0 To m_listbox.ListCount - 1
        If m_listbox.selected(count) Then
            list.Add m_list.item(count + 1)
        End If
    Next count
End Sub

Private Property Get IWindow_hWnd() As Long
    IWindow_hWnd = UserControl.hwnd
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Function drawItem(item As DRAWITEMSTRUCT) As Long
    If item.itemAction = ODA_FOCUS Then
        If item.itemState And ODS_FOCUS Then
            DrawFocusRect item.hdc, item.rcItem
            drawItem = 1
            Exit Function
        End If
    End If

    Dim listItem As CNicklistItem
    
    If item.itemID = -1 Then
        Exit Function
    End If
    
    Set listItem = m_list.item(item.itemID + 1)
    
    Dim fillBrush As Long
    Dim fillColour As Long
    Dim windowColour As Long
    Dim userStyle As CUserStyle
    
    Set userStyle = m_session.getUserStyle(listItem.text, listItem.prefix)
    
    If item.itemState And ODS_SELECTED Then
        fillBrush = GetSysColorBrush(COLOR_HIGHLIGHT)
        SetBkColor item.hdc, GetSysColor(COLOR_HIGHLIGHT)
        SetTextColor item.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
        FillRect item.hdc, item.rcItem, fillBrush
    Else
        fillBrush = CreateSolidBrush(getPaletteEntry(g_nicklistBack))
        SetBkColor item.hdc, getPaletteEntry(g_nicklistBack)
        
        If Not userStyle Is Nothing Then
            SetTextColor item.hdc, colourThemes.currentTheme.paletteEntry(userStyle.foreColour)
        Else
            SetTextColor item.hdc, getPaletteEntry(g_nicklistFore)
        End If
        
        FillRect item.hdc, item.rcItem, fillBrush
        DeleteObject fillBrush
    End If
    
    Dim textRect As RECT
    
    textRect = item.rcItem
    
    Dim oldFont As Long
    Dim newFont As Long
    
    If settings.setting("boldNicks", estBoolean) Then
        newFont = m_fontManager.getFont(True, False, False)
    Else
        newFont = m_fontManager.getDefaultFont
    End If
    
    If newFont <> 0 Then
        oldFont = SelectObject(item.hdc, newFont)
    End If
    
    Dim drewImage As Boolean
    
    If Not userStyle Is Nothing Then
        If Not userStyle.image Is Nothing And settings.setting("nicknameIcons", estBoolean) Then
            userStyle.image.draw item.hdc, textRect.left, textRect.top, m_fontManager.fontHeight, _
                m_fontManager.fontHeight
            textRect.left = textRect.left + m_fontManager.fontHeight + 1
            drewImage = True
        End If
    End If
    
    If drewImage Then
        swiftTextOut item.hdc, textRect.left, textRect.top, 0, VarPtr(textRect), listItem.text
    Else
        swiftTextOut item.hdc, textRect.left + 5, textRect.top, 0, VarPtr(textRect), listItem.prefix & _
            listItem.text
    End If
        
    drawItem = True
    
    If oldFont <> 0 Then
        SelectObject item.hdc, oldFont
    End If
End Function

Public Sub addItem(text As String, prefix As String)
    If m_list.count < 1 Then
        realAddItem text, prefix, 1
        Exit Sub
    End If

    Dim insertValue As String
    
    insertValue = LCase(text)
    
    Dim pivotIndex As Long
    Dim pivotText As String
    Dim pivotMin As Long
    Dim pivotMax As Long
    Dim insertPos As Long
    Dim pivotPrefix As String
    Dim compResult As Integer
    
    pivotMin = 1
    pivotMax = m_list.count
    
    Do While (pivotMax - pivotMin) > 0
        pivotIndex = (((pivotMax - pivotMin)) / 2) + pivotMin
        pivotText = LCase(m_list.item(pivotIndex).text)
        pivotPrefix = m_list.item(pivotIndex).prefix
        
        compResult = m_session.comparePrefix(prefix, pivotPrefix)
        
        If compResult = 0 Then
            If insertValue < pivotText Then
                pivotMax = pivotIndex - 1
            Else
                pivotMin = pivotIndex + 1
            End If
        ElseIf compResult = -1 Then
            pivotMax = pivotIndex - 1
        Else
            pivotMin = pivotIndex + 1
        End If
    Loop
    
    pivotText = LCase(m_list.item(pivotMax).text)
    
    compResult = m_session.comparePrefix(prefix, m_list.item(pivotMax).prefix)
    
    If compResult = 0 Then
        If insertValue < pivotText Then
            insertPos = pivotMax
        Else
            insertPos = pivotMax + 1
        End If
    ElseIf compResult = -1 Then
        insertPos = pivotMax
    Else
        insertPos = pivotMax + 1
    End If
    
    realAddItem text, prefix, insertPos
End Sub

Public Sub removeItem(text As String, prefix As String)
    If m_list.count = 0 Then
        Exit Sub
    End If
    
    Dim removeText As String
    Dim pivotIndex As Long
    Dim pivotText As String
    Dim pivotPrefix As String
    Dim pivotMin As Long
    Dim pivotMax As Long
    Dim compResult As Integer
    
    removeText = LCase(text)
    
    pivotMin = 1
    pivotMax = m_list.count
    
    Do While (pivotMax - pivotMin) > 0
        pivotIndex = (((pivotMax - pivotMin)) / 2) + pivotMin
        pivotText = LCase(m_list.item(pivotIndex).text)
        pivotPrefix = m_list.item(pivotIndex).prefix
        
        compResult = m_session.comparePrefix(prefix, pivotPrefix)
        
        If compResult = 0 Then
            If removeText = pivotText Then
                m_list.Remove pivotIndex
                m_listbox.removeItem pivotIndex - 1
                Exit Sub
            ElseIf removeText < pivotText Then
                pivotMax = pivotIndex - 1
            Else
                pivotMin = pivotIndex + 1
            End If
        ElseIf compResult = -1 Then
            pivotMax = pivotIndex - 1
        Else
            pivotMin = pivotIndex + 1
        End If
    Loop
    
    pivotText = LCase(m_list.item(pivotMax).text)
    
    If pivotText = removeText Then
        m_list.Remove pivotMax
        m_listbox.removeItem pivotMax - 1
    End If
End Sub

Public Sub clearItems()
    m_list.clear
    m_listbox.clear
End Sub

Private Sub realAddItem(text As String, prefix As String, index As Long)
    Dim newItem As New CNicklistItem
    
    newItem.text = text
    newItem.prefix = prefix
    
    m_list.Add newItem, index
    m_listbox.addItem newItem.text, index - 1
End Sub

Private Sub UserControl_Initialize()
    m_itemHeight = 1

    g_hook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf listBoxHook, App.hInstance, App.ThreadID)
    Set m_listbox = Controls.Add("VB.ListBox", "listBox")
    UnhookWindowsHookEx g_hook
    
    m_listbox.Appearance = 0
    
    m_listbox.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    m_listbox.visible = True
    
    initMessages
End Sub

Public Sub refresh()
    RedrawWindow m_listbox.hwnd, ByVal 0, ByVal 0, RDW_INVALIDATE
End Sub

Private Sub initMessages()
    AttachMessage Me, m_listbox.hwnd, WM_LBUTTONDOWN
    AttachMessage Me, m_listbox.hwnd, WM_LBUTTONUP
    AttachMessage Me, m_listbox.hwnd, WM_LBUTTONDBLCLK
    AttachMessage Me, m_listbox.hwnd, WM_MOUSEMOVE
    AttachMessage Me, m_listbox.hwnd, WM_RBUTTONDOWN
    AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
    AttachMessage Me, UserControl.hwnd, WM_MEASUREITEM
    AttachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub deInitMessages()
    DetachMessage Me, m_listbox.hwnd, WM_LBUTTONDOWN
    DetachMessage Me, m_listbox.hwnd, WM_LBUTTONUP
    DetachMessage Me, m_listbox.hwnd, WM_LBUTTONDBLCLK
    DetachMessage Me, m_listbox.hwnd, WM_MOUSEMOVE
    DetachMessage Me, m_listbox.hwnd, WM_RBUTTONDOWN
    DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
    DetachMessage Me, UserControl.hwnd, WM_MEASUREITEM
    DetachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub UserControl_Resize()
    MoveWindow m_listbox.hwnd, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
End Sub

Private Sub UserControl_Terminate()
    deInitMessages
End Sub
