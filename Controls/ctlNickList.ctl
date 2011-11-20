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
10        Set item = m_list.item(index)
End Property

Public Property Get itemCount() As Long
10        itemCount = m_list.count
End Property

Public Property Get session() As CSession
10        Set session = m_session
End Property

Public Property Let session(newValue As CSession)
10        Set m_session = newValue
End Property

Private Sub IColourUser_coloursUpdated()
10        refresh
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
10        fontsUpdated
End Sub

Private Sub fontsUpdated()
10        m_itemHeight = m_fontManager.fontHeight
          
20        If Not m_listbox Is Nothing Then
30            SendMessage m_listbox.hwnd, LB_SETITEMHEIGHT, 0, ByVal m_itemHeight
              
40            If m_list.count > 0 Then
                  Dim TopIndex As Long
                  
50                TopIndex = m_listbox.TopIndex
                  
60                If TopIndex > (m_list.count - 1) - (UserControl.ScaleHeight) / m_itemHeight Then
                      'If the font change has left space at the bottom of
                      'the listbox, update the scroll index appropriately.
70                    m_listbox.TopIndex = m_list.count - 1
80                End If
90            End If
              
100           Me.refresh
110       End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        Select Case CurrentMessage
              Case WM_DRAWITEM
20                ISubclass_MsgResponse = emrConsume
30            Case WM_MOUSEMOVE
40                ISubclass_MsgResponse = emrPreprocess
50            Case WM_LBUTTONDOWN
60                ISubclass_MsgResponse = emrConsume
70            Case Else
80                ISubclass_MsgResponse = emrConsume
90        End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
          Dim count As Long
          Dim coordY As Long
          Dim coordX As Long
          Dim TopIndex As Long
          Dim itemIndex As Long

10        Select Case iMsg
              Case WM_MOUSEMOVE
20                coordX = LoWord(lParam)
              
30                If coordX <= 0 Or coordX > 5 Then
40                    If m_resizing = False Then
50                        m_listbox.MousePointer = vbArrow
60                    End If
70                ElseIf coordX <= 5 Then
80                    m_listbox.MousePointer = vbSizeWE
90                End If
                  
100               If m_resizing = True Then
110                   RaiseEvent changeWidth(UserControl.ScaleWidth - coordX)
120               End If
130           Case WM_LBUTTONDOWN
140               coordX = LoWord(lParam)
              
150               If coordX > 0 And coordX <= 5 Then
160                   m_resizing = True
170                   SetCapture m_listbox.hwnd
180               Else
190                   CallOldWindowProc hwnd, iMsg, wParam, lParam
200               End If
210           Case WM_LBUTTONUP
220               ReleaseCapture
230               m_resizing = False
240           Case WM_LBUTTONDBLCLK
250               coordY = HiWord(lParam)
                  
260               TopIndex = m_listbox.TopIndex
270               itemIndex = TopIndex + Fix(coordY / m_itemHeight)
                  
280               If itemIndex < m_listbox.ListCount Then
290                   RaiseEvent doubleClicked(m_list.item(itemIndex + 1))
300               End If
310           Case WM_RBUTTONDOWN
320               coordY = HiWord(lParam)
                  
330               TopIndex = m_listbox.TopIndex
340               itemIndex = TopIndex + Fix(coordY / m_itemHeight)
                  
350               If itemIndex < m_listbox.ListCount Then
                      Dim selected As Long
                      
360                   selected = SendMessage(m_listbox.hwnd, LB_GETSEL, itemIndex, ByVal 0&)
                      
                      'If the user right clicked on an unselected
                      'item, we'll deselect any other items and select
                      'only the clicked item.  Otherwise, we reselect
                      'the clicked item to give it the focus caret.
                      
370                   If selected = 0 Then
380                       SendMessage m_listbox.hwnd, LB_SETSEL, 0, ByVal -1
390                       SendMessage m_listbox.hwnd, LB_SETSEL, 1, ByVal itemIndex
400                   ElseIf selected > 0 Then
410                       SendMessage m_listbox.hwnd, LB_SETSEL, 1, ByVal itemIndex
420                   End If
430               End If
                  
                  Dim firstSelected As Long
                  
440               For count = 0 To m_listbox.ListCount - 1
450                   If m_listbox.selected(count) Then
460                       firstSelected = count
470                       Exit For
480                   End If
490               Next count
                  
500               RaiseEvent rightClicked(m_list.item(firstSelected + 1))
510           Case WM_DRAWITEM
                  Dim item As DRAWITEMSTRUCT
                  
520               CopyMemory item, ByVal lParam, Len(item)
530               ISubclass_WindowProc = drawItem(item)
540           Case WM_MEASUREITEM
                  Dim measureItem As MEASUREITEMSTRUCT
                  
550               CopyMemory measureItem, ByVal lParam, Len(measureItem)
560               measureItem.itemHeight = m_itemHeight
570               CopyMemory ByVal lParam, measureItem, Len(measureItem)
              
580               ISubclass_WindowProc = True
590           Case WM_CTLCOLORLISTBOX
600               If m_fillbrush <> 0 Then
610                   DeleteObject m_fillbrush
620               End If
              
630               m_fillbrush = CreateSolidBrush(getPaletteEntry(g_nicklistBack))
                  
640               ISubclass_WindowProc = m_fillbrush
650       End Select
End Function

Public Sub getSelectedItems(list As cArrayList)
          Dim count As Long
          
10        For count = 0 To m_listbox.ListCount - 1
20            If m_listbox.selected(count) Then
30                list.Add m_list.item(count + 1)
40            End If
50        Next count
End Sub

Private Property Get IWindow_hWnd() As Long
10        IWindow_hWnd = UserControl.hwnd
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Function drawItem(item As DRAWITEMSTRUCT) As Long
10        If item.itemAction = ODA_FOCUS Then
20            If item.itemState And ODS_FOCUS Then
30                DrawFocusRect item.hdc, item.rcItem
40                drawItem = 1
50                Exit Function
60            End If
70        End If

          Dim listItem As CNicklistItem
          
80        If item.itemID = -1 Then
90            Exit Function
100       End If
          
110       Set listItem = m_list.item(item.itemID + 1)
          
          Dim fillBrush As Long
          Dim fillColour As Long
          Dim windowColour As Long
          Dim userStyle As CUserStyle
          
120       Set userStyle = m_session.getUserStyle(listItem.text, listItem.prefix)
          
130       If item.itemState And ODS_SELECTED Then
140           fillBrush = GetSysColorBrush(COLOR_HIGHLIGHT)
150           SetBkColor item.hdc, GetSysColor(COLOR_HIGHLIGHT)
160           SetTextColor item.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
170           FillRect item.hdc, item.rcItem, fillBrush
180       Else
190           fillBrush = CreateSolidBrush(getPaletteEntry(g_nicklistBack))
200           SetBkColor item.hdc, getPaletteEntry(g_nicklistBack)
              
210           If Not userStyle Is Nothing Then
220               SetTextColor item.hdc, colourThemes.currentTheme.paletteEntry(userStyle.foreColour)
230           Else
240               SetTextColor item.hdc, getPaletteEntry(g_nicklistFore)
250           End If
              
260           FillRect item.hdc, item.rcItem, fillBrush
270           DeleteObject fillBrush
280       End If
          
          Dim textRect As RECT
          
290       textRect = item.rcItem
          
          Dim oldFont As Long
          Dim newFont As Long
          
300       If settings.setting("boldNicks", estBoolean) Then
310           newFont = m_fontManager.getFont(True, False, False)
320       Else
330           newFont = m_fontManager.getDefaultFont
340       End If
          
350       If newFont <> 0 Then
360           oldFont = SelectObject(item.hdc, newFont)
370       End If
          
          Dim drewImage As Boolean
          
380       If Not userStyle Is Nothing Then
390           If Not userStyle.image Is Nothing And settings.setting("nicknameIcons", estBoolean) Then
400               userStyle.image.draw item.hdc, textRect.left, textRect.top, m_fontManager.fontHeight, _
                      m_fontManager.fontHeight
410               textRect.left = textRect.left + m_fontManager.fontHeight + 1
420               drewImage = True
430           End If
440       End If
          
450       If drewImage Then
460           swiftTextOut item.hdc, textRect.left, textRect.top, 0, VarPtr(textRect), listItem.text
470       Else
480           swiftTextOut item.hdc, textRect.left + 5, textRect.top, 0, VarPtr(textRect), listItem.prefix & _
                  listItem.text
490       End If
              
500       drawItem = True
          
510       If oldFont <> 0 Then
520           SelectObject item.hdc, oldFont
530       End If
End Function

Public Sub addItem(text As String, prefix As String)
10        If m_list.count < 1 Then
20            realAddItem text, prefix, 1
30            Exit Sub
40        End If

          Dim insertValue As String
          
50        insertValue = LCase(text)
          
          Dim pivotIndex As Long
          Dim pivotText As String
          Dim pivotMin As Long
          Dim pivotMax As Long
          Dim insertPos As Long
          Dim pivotPrefix As String
          Dim compResult As Integer
          
60        pivotMin = 1
70        pivotMax = m_list.count
          
80        Do While (pivotMax - pivotMin) > 0
90            pivotIndex = (((pivotMax - pivotMin)) / 2) + pivotMin
100           pivotText = LCase(m_list.item(pivotIndex).text)
110           pivotPrefix = m_list.item(pivotIndex).prefix
              
120           compResult = m_session.comparePrefix(prefix, pivotPrefix)
              
130           If compResult = 0 Then
140               If insertValue < pivotText Then
150                   pivotMax = pivotIndex - 1
160               Else
170                   pivotMin = pivotIndex + 1
180               End If
190           ElseIf compResult = -1 Then
200               pivotMax = pivotIndex - 1
210           Else
220               pivotMin = pivotIndex + 1
230           End If
240       Loop
          
250       pivotText = LCase(m_list.item(pivotMax).text)
          
260       compResult = m_session.comparePrefix(prefix, m_list.item(pivotMax).prefix)
          
270       If compResult = 0 Then
280           If insertValue < pivotText Then
290               insertPos = pivotMax
300           Else
310               insertPos = pivotMax + 1
320           End If
330       ElseIf compResult = -1 Then
340           insertPos = pivotMax
350       Else
360           insertPos = pivotMax + 1
370       End If
          
380       realAddItem text, prefix, insertPos
End Sub

Public Sub removeItem(text As String, prefix As String)
10        If m_list.count = 0 Then
20            Exit Sub
30        End If
          
          Dim removeText As String
          Dim pivotIndex As Long
          Dim pivotText As String
          Dim pivotPrefix As String
          Dim pivotMin As Long
          Dim pivotMax As Long
          Dim compResult As Integer
          
40        removeText = LCase(text)
          
50        pivotMin = 1
60        pivotMax = m_list.count
          
70        Do While (pivotMax - pivotMin) > 0
80            pivotIndex = (((pivotMax - pivotMin)) / 2) + pivotMin
90            pivotText = LCase(m_list.item(pivotIndex).text)
100           pivotPrefix = m_list.item(pivotIndex).prefix
              
110           compResult = m_session.comparePrefix(prefix, pivotPrefix)
              
120           If compResult = 0 Then
130               If removeText = pivotText Then
140                   m_list.Remove pivotIndex
150                   m_listbox.removeItem pivotIndex - 1
160                   Exit Sub
170               ElseIf removeText < pivotText Then
180                   pivotMax = pivotIndex - 1
190               Else
200                   pivotMin = pivotIndex + 1
210               End If
220           ElseIf compResult = -1 Then
230               pivotMax = pivotIndex - 1
240           Else
250               pivotMin = pivotIndex + 1
260           End If
270       Loop
          
280       pivotText = LCase(m_list.item(pivotMax).text)
          
290       If pivotText = removeText Then
300           m_list.Remove pivotMax
310           m_listbox.removeItem pivotMax - 1
320       End If
End Sub

Public Sub clearItems()
10        m_list.clear
20        m_listbox.clear
End Sub

Private Sub realAddItem(text As String, prefix As String, index As Long)
          Dim newItem As New CNicklistItem
          
10        newItem.text = text
20        newItem.prefix = prefix
          
30        m_list.Add newItem, index
40        m_listbox.addItem newItem.text, index - 1
End Sub

Private Sub UserControl_Initialize()
10        m_itemHeight = 1

20        g_hook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf listBoxHook, App.hInstance, App.ThreadID)
30        Set m_listbox = Controls.Add("VB.ListBox", "listBox")
40        UnhookWindowsHookEx g_hook
          
50        m_listbox.Appearance = 0
          
60        m_listbox.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
70        m_listbox.visible = True
          
80        initMessages
End Sub

Public Sub refresh()
10        RedrawWindow m_listbox.hwnd, ByVal 0, ByVal 0, RDW_INVALIDATE
End Sub

Private Sub initMessages()
10        AttachMessage Me, m_listbox.hwnd, WM_LBUTTONDOWN
20        AttachMessage Me, m_listbox.hwnd, WM_LBUTTONUP
30        AttachMessage Me, m_listbox.hwnd, WM_LBUTTONDBLCLK
40        AttachMessage Me, m_listbox.hwnd, WM_MOUSEMOVE
50        AttachMessage Me, m_listbox.hwnd, WM_RBUTTONDOWN
60        AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
70        AttachMessage Me, UserControl.hwnd, WM_MEASUREITEM
80        AttachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub deInitMessages()
10        DetachMessage Me, m_listbox.hwnd, WM_LBUTTONDOWN
20        DetachMessage Me, m_listbox.hwnd, WM_LBUTTONUP
30        DetachMessage Me, m_listbox.hwnd, WM_LBUTTONDBLCLK
40        DetachMessage Me, m_listbox.hwnd, WM_MOUSEMOVE
50        DetachMessage Me, m_listbox.hwnd, WM_RBUTTONDOWN
60        DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
70        DetachMessage Me, UserControl.hwnd, WM_MEASUREITEM
80        DetachMessage Me, UserControl.hwnd, WM_CTLCOLORLISTBOX
End Sub

Private Sub UserControl_Resize()
10        MoveWindow m_listbox.hwnd, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
End Sub

Private Sub UserControl_Terminate()
10        deInitMessages
End Sub
