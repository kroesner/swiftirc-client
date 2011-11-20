VERSION 5.00
Begin VB.UserControl ctlTextInput 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.TextBox m_textBox 
      Height          =   285
      Left            =   705
      TabIndex        =   0
      Top             =   750
      Width           =   1275
   End
End
Attribute VB_Name = "ctlTextInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IFontUser
Implements ISubclass
Implements IColourUser

Private m_realWindow As VB.VBControlExtender
Private m_fontManager As CFontManager

Private m_inputHistory As New Collection
Private m_inputHistoryIndex As Long

Public Event tabbed(text As String, start As Long, Length As Long)
Public Event textSubmitted(text As String, ctrl As Boolean)
Public Event mouseWheel(delta As Long)
Public Event pageUp()
Public Event pageDown()

Private m_IPAOHookStruct As IPAOHookStructTextInput
Private Const WM_SETFOCUS = &H7

Private Sub IColourUser_coloursUpdated()
10        updateColours
End Sub

Private Sub updateColours()
10        If g_initialized Then
20            m_textBox.BackColor = colourThemes.currentTheme.paletteEntry(g_textInputBack)
30            m_textBox.ForeColor = colourThemes.currentTheme.paletteEntry(g_textInputFore)
40            UserControl_Paint
50        End If
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
10        If Not m_textBox Is Nothing Then
20            SendMessage m_textBox.hwnd, WM_SETFONT, m_fontManager.getDefaultFont, True
30        End If
          
40        UserControl_Paint
End Sub

Public Function TranslateAccelerator(lpMsg As VBOleGuids.Msg) As Long
10        TranslateAccelerator = S_FALSE
          ' Here you can modify the response to the key down
          ' accelerator command using the values in lpMsg.  This
          ' can be used to capture Tabs, Returns, Arrows etc.
          ' Just process the message as required and return S_OK.
20        If (lpMsg.wParam And &HFFFF&) = vbKeyTab Then
30           Select Case lpMsg.message
             Case WM_KEYDOWN
40              UserControl_KeyDown vbKeyTab, ShiftState
50              TranslateAccelerator = S_OK
60           Case WM_KEYUP
70              UserControl_KeyUp vbKeyTab, ShiftState
80              TranslateAccelerator = S_OK
90           End Select
100       End If
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        Select Case CurrentMessage
              Case WM_PASTE
20                ISubclass_MsgResponse = emrConsume
30            Case Else
40                ISubclass_MsgResponse = emrPreprocess
50        End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
10        Select Case iMsg
              Case WM_MOUSEWHEEL
20                RaiseEvent mouseWheel(HiWord(wParam))
30            Case WM_PASTE
40                processPaste
50            Case WM_SETFOCUS
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
60                Set pOleObject = Me
70                Set pOleInPlaceSite = pOleObject.GetClientSite
80                pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), _
                      VarPtr(ClipRect), VarPtr(FrameInfo)
90                CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
100               pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
110               If Not pOleInPlaceUIWindow Is Nothing Then
120                  pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
130               End If
                  ' Clear up the inbetween implementation:
140               CopyMemory pOleInPlaceActiveObject, 0&, 4
                  ' --------------------------------------------------------------------------
150       End Select
End Function

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Public Property Get text() As String
10        text = m_textBox.text
End Property

Public Property Let text(newValue As String)
10        m_textBox.text = newValue
20        m_textBox.selStart = Len(m_textBox.text)
End Property

Private Sub InsertSymbol(symbol As String)
10        If m_textBox.SelLength > 0 Then
20            m_textBox.SelText = symbol & m_textBox.SelText & symbol
30            m_textBox.selStart = m_textBox.selStart + m_textBox.SelLength
40            m_textBox.SelLength = 0
50        Else
60            m_textBox.SelText = symbol
70        End If
End Sub

Private Sub lineEntered(text As String, ctrl As Boolean)
10        m_inputHistory.Add CVar(text)
20        m_inputHistoryIndex = m_inputHistory.count + 1
          
30        RaiseEvent textSubmitted(UTF8Decode(text), ctrl)
End Sub

Private Sub processPaste()
          Dim text As String
          
10        text = Clipboard.getText
          
          Dim count As Integer
          Dim char As Integer
          Dim lines As New Collection
          
          Dim temp As String

20        For count = 1 To Len(text)
30            char = AscW(Mid$(text, count, 1))
              
40            If char = 10 Or char = 13 Then
50                If count < Len(text) Then
60                    char = AscW(Mid$(text, count + 1, 1))
                      
70                    If char = 10 Or char = 13 Then
80                        count = count + 1
90                    End If
100               End If
                  
110               If LenB(temp) <> 0 Then
120                   lines.Add temp
130               End If
                  
140               temp = vbNullString
150           Else
160               temp = temp & ChrW$(char)
170           End If
180       Next count
          
190       If LenB(temp) <> 0 Then
200           lines.Add temp
210       End If
          
220       If lines.count < 2 Then
230           If lines.count = 0 Then
240               m_textBox.SelText = text
250               m_textBox.selStart = m_textBox.selStart + Len(text)
260           Else
270               m_textBox.SelText = lines.item(1)
280               m_textBox.selStart = m_textBox.selStart + Len(lines.item(1))
290           End If
              
300           Exit Sub
310       End If
          
          Dim line As Variant
          
320       If lines.count > 5 Then
330           If MsgBox("Are you sure you want to paste " & lines.count & " line" & IIf(lines.count = 1, "", _
                  "s") & " of text?", vbQuestion Or vbYesNo, "Pasting") = vbYes Then
                  
340               For Each line In lines
350                   lineEntered CStr(line), False
360               Next line
370           End If
380       Else
390           For Each line In lines
400               lineEntered CStr(line), False
410           Next line
420       End If
End Sub

Private Sub initMessages()
10        AttachMessage Me, m_textBox.hwnd, WM_PASTE
20        AttachMessage Me, m_textBox.hwnd, WM_SETFOCUS
30        AttachMessage Me, m_textBox.hwnd, WM_MOUSEWHEEL
End Sub

Private Sub deInitMessages()
10        DetachMessage Me, m_textBox.hwnd, WM_PASTE
20        DetachMessage Me, m_textBox.hwnd, WM_SETFOCUS
30        DetachMessage Me, m_textBox.hwnd, WM_MOUSEWHEEL
End Sub

Private Sub m_textBox_KeyDown(KeyCode As Integer, Shift As Integer)
10        Select Case KeyCode
              Case vbKeyDown
20                KeyCode = 0
                  
30                If m_inputHistoryIndex < m_inputHistory.count Then
40                    m_inputHistoryIndex = m_inputHistoryIndex + 1
50                    m_textBox.text = m_inputHistory.item(m_inputHistoryIndex)
60                    m_textBox.selStart = Len(m_textBox.text)
70                ElseIf m_inputHistoryIndex = m_inputHistory.count Then
80                    m_textBox.text = vbNullString
90                    m_inputHistoryIndex = m_inputHistoryIndex + 1
100               Else
110                   If LenB(m_textBox.text) <> 0 Then
120                       m_textBox.text = vbNullString
130                   Else
140                       Beep
150                   End If
160               End If
170           Case vbKeyUp
180               KeyCode = 0
              
190               If m_inputHistory.count < 1 Or m_inputHistoryIndex = 1 Then
200                   Beep
210                   Exit Sub
220               End If
                  
230               m_inputHistoryIndex = m_inputHistoryIndex - 1
240               m_textBox.text = m_inputHistory.item(m_inputHistoryIndex)
250               m_textBox.selStart = Len(m_textBox.text)
260           Case vbKeyPageUp
270               RaiseEvent pageUp
280           Case vbKeyPageDown
290               RaiseEvent pageDown
300           Case vbKeyReturn
310               If LenB(m_textBox.text) <> 0 Then
320                   If Shift And vbCtrlMask Then
330                       lineEntered m_textBox.text, True
340                   Else
350                       lineEntered m_textBox.text, False
360                   End If
                      
370                   m_textBox.text = vbNullString
380               End If
390       End Select
End Sub

Public Sub replaceText(ByVal start As Long, ByVal Length As Long, newText As String)
10        LockWindowUpdate m_textBox.hwnd

20        m_textBox.selStart = start
30        m_textBox.SelLength = Length
40        m_textBox.SelText = newText
50        m_textBox.SelLength = 0
60        m_textBox.selStart = start + Len(newText)
          
70        LockWindowUpdate 0
80        m_textBox.refresh
End Sub

Private Sub m_textbox_KeyPress(KeyAscii As Integer)
10        Select Case KeyAscii
              Case 10
20                KeyAscii = 0
30            Case vbKeyReturn
40                KeyAscii = 0
50            Case 2
60                InsertSymbol Chr$(2)
70                KeyAscii = 0
80            Case 11
90                InsertSymbol Chr$(3)
100               KeyAscii = 0
110           Case 20
120               InsertSymbol Chr$(4)
130               KeyAscii = 0
140           Case 21
150               InsertSymbol Chr$(31)
160               KeyAscii = 0
170           Case 18
180               InsertSymbol Chr$(22)
190               KeyAscii = 0
200           Case 15
210               InsertSymbol Chr$(15)
220               KeyAscii = 0
230       End Select
End Sub

Private Property Get ShiftState() As Integer
10        If GetAsyncKeyState(vbKeyShift) <> 0 Then
20            ShiftState = vbShiftMask
30        End If

40        If GetAsyncKeyState(vbKeyControl) <> 0 Then
50            ShiftState = ShiftState Or vbCtrlMask
60        End If
End Property

Private Sub UserControl_Initialize()
          'Set m_textBox = Controls.Add("VB.TextBox", "txtInput")
10        m_textBox.visible = True
20        initMessages
          
30        m_textBox.Appearance = 0
40        m_textBox.BorderStyle = 0
          
          Dim IPAO As IOleInPlaceActiveObject

50        With m_IPAOHookStruct
60           Set IPAO = Me
70           CopyMemory .IPAOReal, IPAO, 4
80           CopyMemory .TBEx, Me, 4
90           .lpVTable = IPAOVTableTextInput
100          .ThisPointer = VarPtr(m_IPAOHookStruct)
110       End With
          
120       updateColours
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
10        Select Case KeyCode
              Case vbKeyTab
20                KeyCode = 0
                  
                  Dim textLen As Long
                  
30                textLen = Len(m_textBox.text)
                  
40                If textLen = 0 Then
50                    Exit Sub
60                ElseIf textLen = 1 Then
70                    RaiseEvent tabbed(left(m_textBox.text, 1), 0, 1)
80                    Exit Sub
90                End If
              
                  Dim tabStart As Long
                  Dim tabEnd As Long
                  
100               If m_textBox.selStart = textLen Then
110                   tabEnd = textLen + 1
120               Else
130                   If m_textBox.selStart = 0 Then
140                       tabEnd = InStr(m_textBox.text, Chr$(32))
150                   Else
160                       tabEnd = InStr(m_textBox.selStart + 1, m_textBox.text, Chr$(32))
170                   End If
                      
180                   If tabEnd = 0 Then
190                       tabEnd = textLen + 1
200                   End If
210               End If
                  
220               tabStart = InStrRev(m_textBox.text, Chr$(32), tabEnd - 1)
                  
230               If tabEnd - tabStart > 1 Then
240                   RaiseEvent tabbed(Mid$(m_textBox.text, tabStart + 1, (tabEnd - tabStart) - 1), _
                          tabStart, (tabEnd - tabStart) - 1)
250               End If
260       End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub UserControl_Paint()
          Dim brush As Long
          
10        brush = CreateSolidBrush(colourThemes.currentTheme.paletteEntry(g_textInputFore))

20        FrameRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              brush
              
30        DeleteObject brush
End Sub

Private Sub UserControl_Resize()
10        m_textBox.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
End Sub

Private Sub UserControl_Terminate()
10        With m_IPAOHookStruct
20          CopyMemory .IPAOReal, 0&, 4
30          CopyMemory .TBEx, 0&, 4
40        End With
         
50        deInitMessages
End Sub
