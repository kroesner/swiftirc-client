VERSION 5.00
Begin VB.UserControl ctlButton 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
End
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements ISubclass
Implements IColourUser

Private m_state As eButtonState
Private m_caption As String
Private m_hasFocus As Boolean

Private Enum eButtonState
    bsNormal
    bsMouseover
    bsMouseDown
    bsSelected
End Enum

Public Event clicked()

Public Property Get visible() As Boolean
    visible = UserControl.Extender.visible
End Property

Public Property Let visible(newValue As Boolean)
    UserControl.Extender.visible = newValue
End Property


Private Sub IColourUser_coloursUpdated()
    UserControl_Paint
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_LBUTTONDBLCLK
            ISubclass_MsgResponse = emrConsume
        Case Else
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
        Case WM_LBUTTONDBLCLK
            UserControl_MouseDown vbKeyLButton, 0, 0, 0
            ISubclass_WindowProc = 1
    End Select
End Function

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = UserControl.Extender
End Property

Private Property Let IWindow_realWindow(RHS As Object)
End Property

Public Property Get caption() As String
    caption = m_caption
End Property

Public Property Let caption(newValue As String)
    m_caption = newValue
    parsePrefix newValue
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        RaiseEvent clicked
    End If
End Sub

Private Sub UserControl_GotFocus()
    m_hasFocus = True
    UserControl_Paint
End Sub

Private Sub initMessages()
    AttachMessage Me, UserControl.hwnd, WM_LBUTTONDBLCLK
End Sub

Private Sub deInitMessages()
    DetachMessage Me, UserControl.hwnd, WM_LBUTTONDBLCLK
End Sub

Private Sub UserControl_Initialize()
    initMessages
    SelectObject UserControl.hdc, g_fontUI
End Sub

Private Sub UserControl_Terminate()
    deInitMessages
End Sub

Private Sub UserControl_LostFocus()
    m_hasFocus = False
    UserControl_Paint
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        RaiseEvent clicked
    ElseIf KeyCode = vbKeySpace Then
        UserControl_MouseDown vbKeyLButton, 0, 0, 0
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        UserControl_MouseUp vbKeyLButton, 0, -1, -1
        RaiseEvent clicked
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbKeyLButton Then
        m_state = bsMouseDown
        UserControl_Paint
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not m_state = bsMouseDown And Not m_state = bsMouseover Then
        m_state = bsMouseover
        SetCapture UserControl.hwnd
        UserControl_Paint
    Else
        If m_state = bsMouseover Then
            If x < 0 Or y < 0 Or x > UserControl.ScaleWidth Or y > UserControl.ScaleHeight Then
                ReleaseCapture
                m_state = bsNormal
                UserControl_Paint
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCapture UserControl.hwnd

    If Button <> vbKeyLButton Then
        Exit Sub
    End If

    If x < 0 Or y < 0 Or x > UserControl.ScaleWidth Or y > UserControl.ScaleHeight Then
        m_state = bsNormal
        ReleaseCapture
        UserControl_Paint
    Else
        m_state = bsNormal
        UserControl_Paint
        ReleaseCapture
        RaiseEvent clicked
    End If
End Sub

Private Sub parsePrefix(caption As String)
    Dim count As Long
    
    For count = 1 To Len(caption)
        If Mid$(caption, count, 1) = "&" Then
            If count < Len(caption) Then
                If Mid$(caption, count + 1, 1) <> "&" Then
                    UserControl.AccessKeys = Mid$(caption, count + 1, 1)
                End If
            End If
        End If
    Next count
End Sub

Private Sub UserControl_Paint()
    If Not g_initialized Then
        Exit Sub
    End If
    
    Dim oldPen As Long
    Dim oldBrush As Long

    oldBrush = SelectObject(UserControl.hdc, colourManager.getBrush(SWIFTCOLOUR_CONTROLBACK))
    oldPen = SelectObject(UserControl.hdc, colourManager.getPen(SWIFTPEN_BORDER))
    
    Rectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight

    Dim textRect As RECT
    
    textRect.left = 3
    textRect.top = 3
    textRect.right = UserControl.ScaleWidth - 3
    textRect.bottom = UserControl.ScaleHeight - 3
    
    Dim focusRect As RECT
    
    focusRect.left = 2
    focusRect.top = 2
    focusRect.right = UserControl.ScaleWidth - 2
    focusRect.bottom = UserControl.ScaleHeight - 2
    
    If m_hasFocus Then
        SetTextColor UserControl.hdc, Not colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
        DrawFocusRect UserControl.hdc, focusRect
    End If
    
    If m_state = bsNormal Then
        SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
    ElseIf m_state = bsMouseover Then
        SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
    ElseIf m_state = bsMouseDown Then
        textRect.left = 7
        textRect.top = 7
        SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
    End If
    
    swiftDrawText UserControl.hdc, m_caption, VarPtr(textRect), DT_SINGLELINE Or DT_VCENTER Or DT_CENTER Or DT_END_ELLIPSIS
        
    SelectObject UserControl.hdc, oldBrush
    SelectObject UserControl.hdc, oldPen
End Sub
