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
Private m_fontmanager As CFontManager

Private m_inputHistory As New Collection
Private m_inputHistoryIndex As Long

Public Event tabbed(text As String, start As Long, length As Long)
Public Event textSubmitted(text As String, ctrl As Boolean)
Public Event mouseWheel(delta As Long)
Public Event pageUp()
Public Event pageDown()

Private m_IPAOHookStruct As IPAOHookStructTextInput
Private Const WM_SETFOCUS = &H7

Private Sub IColourUser_coloursUpdated()
    updateColours
End Sub

Private Sub updateColours()
    If g_initialized Then
        m_textBox.BackColor = colourThemes.currentTheme.paletteEntry(g_textInputBack)
        m_textBox.ForeColor = colourThemes.currentTheme.paletteEntry(g_textInputFore)
        UserControl_Paint
    End If
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontmanager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
    If Not m_textBox Is Nothing Then
        SendMessage m_textBox.hwnd, WM_SETFONT, m_fontmanager.getDefaultFont, True
    End If
    
    UserControl_Paint
End Sub

Public Function TranslateAccelerator(lpMsg As VBOleGuids.Msg) As Long
    TranslateAccelerator = S_FALSE
    ' Here you can modify the response to the key down
    ' accelerator command using the values in lpMsg.  This
    ' can be used to capture Tabs, Returns, Arrows etc.
    ' Just process the message as required and return S_OK.
    If (lpMsg.wParam And &HFFFF&) = vbKeyTab Then
       Select Case lpMsg.message
       Case WM_KEYDOWN
          UserControl_KeyDown vbKeyTab, ShiftState
          TranslateAccelerator = S_OK
       Case WM_KEYUP
          UserControl_KeyUp vbKeyTab, ShiftState
          TranslateAccelerator = S_OK
       End Select
    End If
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_PASTE
            ISubclass_MsgResponse = emrConsume
        Case Else
            ISubclass_MsgResponse = emrPreprocess
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
        Case WM_MOUSEWHEEL
            RaiseEvent mouseWheel(HiWord(wParam))
        Case WM_PASTE
            processPaste
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
            pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
            CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
            pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
            If Not pOleInPlaceUIWindow Is Nothing Then
               pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
            End If
            ' Clear up the inbetween implementation:
            CopyMemory pOleInPlaceActiveObject, 0&, 4
            ' --------------------------------------------------------------------------
    End Select
End Function

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Public Property Get text() As String
    text = m_textBox.text
End Property

Public Property Let text(newValue As String)
    m_textBox.text = newValue
    m_textBox.selStart = Len(m_textBox.text)
End Property

Private Sub InsertSymbol(symbol As String)
    If m_textBox.SelLength > 0 Then
        m_textBox.SelText = symbol & m_textBox.SelText & symbol
        m_textBox.selStart = m_textBox.selStart + m_textBox.SelLength
        m_textBox.SelLength = 0
    Else
        m_textBox.SelText = symbol
    End If
End Sub

Private Sub lineEntered(text As String, ctrl As Boolean)
    m_inputHistory.Add CVar(text)
    m_inputHistoryIndex = m_inputHistory.count + 1
    
    RaiseEvent textSubmitted(UTF8Decode(text), ctrl)
End Sub

Private Sub processPaste()
    Dim text As String
    
    text = Clipboard.getText
    
    Dim count As Integer
    Dim char As Integer
    Dim lines As New Collection
    
    Dim temp As String

    For count = 1 To Len(text)
        char = AscW(Mid$(text, count, 1))
        
        If char = 10 Or char = 13 Then
            If count < Len(text) Then
                char = AscW(Mid$(text, count + 1, 1))
                
                If char = 10 Or char = 13 Then
                    count = count + 1
                End If
            End If
            
            If LenB(temp) <> 0 Then
                lines.Add temp
            End If
            
            temp = vbNullString
        Else
            temp = temp & ChrW$(char)
        End If
    Next count
    
    If LenB(temp) <> 0 Then
        lines.Add temp
    End If
    
    If lines.count < 2 Then
        If lines.count = 0 Then
            m_textBox.SelText = text
            m_textBox.selStart = m_textBox.selStart + Len(text)
        Else
            m_textBox.SelText = lines.item(1)
            m_textBox.selStart = m_textBox.selStart + Len(lines.item(1))
        End If
        
        Exit Sub
    End If
    
    Dim line As Variant
    
    If lines.count > 5 Then
        If MsgBox("Are you sure you want to paste " & lines.count & " line" & IIf(lines.count = 1, "", "s") & " of text?", vbQuestion Or vbYesNo, "Pasting") = vbYes Then
            
            For Each line In lines
                lineEntered CStr(line), False
            Next line
        End If
    Else
        For Each line In lines
            lineEntered CStr(line), False
        Next line
    End If
End Sub

Private Sub initMessages()
    AttachMessage Me, m_textBox.hwnd, WM_PASTE
    AttachMessage Me, m_textBox.hwnd, WM_SETFOCUS
    AttachMessage Me, m_textBox.hwnd, WM_MOUSEWHEEL
End Sub

Private Sub deInitMessages()
    DetachMessage Me, m_textBox.hwnd, WM_PASTE
    DetachMessage Me, m_textBox.hwnd, WM_SETFOCUS
    DetachMessage Me, m_textBox.hwnd, WM_MOUSEWHEEL
End Sub

Private Sub m_textBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            KeyCode = 0
            
            If m_inputHistoryIndex < m_inputHistory.count Then
                m_inputHistoryIndex = m_inputHistoryIndex + 1
                m_textBox.text = m_inputHistory.item(m_inputHistoryIndex)
                m_textBox.selStart = Len(m_textBox.text)
            ElseIf m_inputHistoryIndex = m_inputHistory.count Then
                m_textBox.text = vbNullString
                m_inputHistoryIndex = m_inputHistoryIndex + 1
            Else
                If LenB(m_textBox.text) <> 0 Then
                    m_textBox.text = vbNullString
                Else
                    Beep
                End If
            End If
        Case vbKeyUp
            KeyCode = 0
        
            If m_inputHistory.count < 1 Or m_inputHistoryIndex = 1 Then
                Beep
                Exit Sub
            End If
            
            m_inputHistoryIndex = m_inputHistoryIndex - 1
            m_textBox.text = m_inputHistory.item(m_inputHistoryIndex)
            m_textBox.selStart = Len(m_textBox.text)
        Case vbKeyPageUp
            RaiseEvent pageUp
        Case vbKeyPageDown
            RaiseEvent pageDown
        Case vbKeyReturn
            If LenB(m_textBox.text) <> 0 Then
                If Shift And vbCtrlMask Then
                    lineEntered m_textBox.text, True
                Else
                    lineEntered m_textBox.text, False
                End If
                
                m_textBox.text = vbNullString
            End If
    End Select
End Sub

Public Sub replaceText(ByVal start As Long, ByVal length As Long, newText As String)
    LockWindowUpdate m_textBox.hwnd

    m_textBox.selStart = start
    m_textBox.SelLength = length
    m_textBox.SelText = newText
    m_textBox.SelLength = 0
    m_textBox.selStart = start + Len(newText)
    
    LockWindowUpdate 0
    m_textBox.refresh
End Sub

Private Sub m_textbox_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 10
            KeyAscii = 0
        Case vbKeyReturn
            KeyAscii = 0
        Case 2
            InsertSymbol Chr$(2)
            KeyAscii = 0
        Case 11
            InsertSymbol Chr$(3)
            KeyAscii = 0
        Case 20
            InsertSymbol Chr$(4)
            KeyAscii = 0
        Case 21
            InsertSymbol Chr$(31)
            KeyAscii = 0
        Case 18
            InsertSymbol Chr$(22)
            KeyAscii = 0
        Case 15
            InsertSymbol Chr$(15)
            KeyAscii = 0
    End Select
End Sub

Private Property Get ShiftState() As Integer
    If GetAsyncKeyState(vbKeyShift) <> 0 Then
        ShiftState = vbShiftMask
    End If

    If GetAsyncKeyState(vbKeyControl) <> 0 Then
        ShiftState = ShiftState Or vbCtrlMask
    End If
End Property

Private Sub UserControl_Initialize()
    'Set m_textBox = Controls.Add("VB.TextBox", "txtInput")
    m_textBox.visible = True
    initMessages
    
    m_textBox.Appearance = 0
    m_textBox.BorderStyle = 0
    
    Dim IPAO As IOleInPlaceActiveObject

    With m_IPAOHookStruct
       Set IPAO = Me
       CopyMemory .IPAOReal, IPAO, 4
       CopyMemory .TBEx, Me, 4
       .lpVTable = IPAOVTableTextInput
       .ThisPointer = VarPtr(m_IPAOHookStruct)
    End With
    
    updateColours
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab
            KeyCode = 0
            
            Dim textLen As Long
            
            textLen = Len(m_textBox.text)
            
            If textLen = 0 Then
                Exit Sub
            ElseIf textLen = 1 Then
                RaiseEvent tabbed(left(m_textBox.text, 1), 0, 1)
                Exit Sub
            End If
        
            Dim tabStart As Long
            Dim tabEnd As Long
            
            If m_textBox.selStart = textLen Then
                tabEnd = textLen + 1
            Else
                If m_textBox.selStart = 0 Then
                    tabEnd = InStr(m_textBox.text, Chr$(32))
                Else
                    tabEnd = InStr(m_textBox.selStart + 1, m_textBox.text, Chr$(32))
                End If
                
                If tabEnd = 0 Then
                    tabEnd = textLen + 1
                End If
            End If
            
            tabStart = InStrRev(m_textBox.text, Chr$(32), tabEnd - 1)
            
            If tabEnd - tabStart > 1 Then
                RaiseEvent tabbed(Mid$(m_textBox.text, tabStart + 1, (tabEnd - tabStart) - 1), tabStart, (tabEnd - tabStart) - 1)
            End If
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub UserControl_Paint()
    Dim brush As Long
    
    brush = CreateSolidBrush(colourThemes.currentTheme.paletteEntry(g_textInputFore))

    FrameRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), brush
        
    DeleteObject brush
End Sub

Private Sub UserControl_Resize()
    m_textBox.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
End Sub

Private Sub UserControl_Terminate()
    With m_IPAOHookStruct
      CopyMemory .IPAOReal, 0&, 4
      CopyMemory .TBEx, 0&, 4
    End With
   
    deInitMessages
End Sub
