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
10        visible = UserControl.Extender.visible
End Property

Public Property Let visible(newValue As Boolean)
10        UserControl.Extender.visible = newValue
End Property


Private Sub IColourUser_coloursUpdated()
10        UserControl_Paint
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
10        Select Case CurrentMessage
              Case WM_LBUTTONDBLCLK
20                ISubclass_MsgResponse = emrConsume
30            Case Else
40        End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
10        Select Case iMsg
              Case WM_LBUTTONDBLCLK
20                UserControl_MouseDown vbKeyLButton, 0, 0, 0
30                ISubclass_WindowProc = 1
40        End Select
End Function

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = UserControl.Extender
End Property

Private Property Let IWindow_realWindow(RHS As Object)
End Property

Public Property Get caption() As String
10        caption = m_caption
End Property

Public Property Let caption(newValue As String)
10        m_caption = newValue
20        parsePrefix newValue
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
10        If KeyAscii = vbKeyReturn Then
20            RaiseEvent clicked
30        End If
End Sub

Private Sub UserControl_GotFocus()
10        m_hasFocus = True
20        UserControl_Paint
End Sub

Private Sub initMessages()
10        AttachMessage Me, UserControl.hwnd, WM_LBUTTONDBLCLK
End Sub

Private Sub deInitMessages()
10        DetachMessage Me, UserControl.hwnd, WM_LBUTTONDBLCLK
End Sub

Private Sub UserControl_Initialize()
10        initMessages
20        SelectObject UserControl.hdc, g_fontUI
End Sub

Private Sub UserControl_Terminate()
10        deInitMessages
20        debugLog "ctlButton terminating: " & m_caption
End Sub

Private Sub UserControl_LostFocus()
10        m_hasFocus = False
20        UserControl_Paint
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeyReturn Then
20            RaiseEvent clicked
30        ElseIf KeyCode = vbKeySpace Then
40            UserControl_MouseDown vbKeyLButton, 0, 0, 0
50        End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeySpace Then
20            UserControl_MouseUp vbKeyLButton, 0, -1, -1
30            RaiseEvent clicked
40        End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If Button = vbKeyLButton Then
20            m_state = bsMouseDown
30            UserControl_Paint
40        End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If Not m_state = bsMouseDown And Not m_state = bsMouseover Then
20            m_state = bsMouseover
30            SetCapture UserControl.hwnd
40            UserControl_Paint
50        Else
60            If m_state = bsMouseover Then
70                If x < 0 Or y < 0 Or x > UserControl.ScaleWidth Or y > UserControl.ScaleHeight Then
80                    ReleaseCapture
90                    m_state = bsNormal
100                   UserControl_Paint
110               End If
120           End If
130       End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
10        SetCapture UserControl.hwnd

20        If Button <> vbKeyLButton Then
30            Exit Sub
40        End If

50        If x < 0 Or y < 0 Or x > UserControl.ScaleWidth Or y > UserControl.ScaleHeight Then
60            m_state = bsNormal
70            ReleaseCapture
80            UserControl_Paint
90        Else
100           m_state = bsNormal
110           UserControl_Paint
120           ReleaseCapture
130           RaiseEvent clicked
140       End If
End Sub

Private Sub parsePrefix(caption As String)
          Dim count As Long
          
10        For count = 1 To Len(caption)
20            If Mid$(caption, count, 1) = "&" Then
30                If count < Len(caption) Then
40                    If Mid$(caption, count + 1, 1) <> "&" Then
50                        UserControl.AccessKeys = Mid$(caption, count + 1, 1)
60                    End If
70                End If
80            End If
90        Next count
End Sub

Private Sub UserControl_Paint()
10        If Not g_initialized Then
20            Exit Sub
30        End If
          
          Dim oldPen As Long
          Dim oldBrush As Long

40        oldBrush = SelectObject(UserControl.hdc, colourManager.getBrush(SWIFTCOLOUR_CONTROLBACK))
50        oldPen = SelectObject(UserControl.hdc, colourManager.getPen(SWIFTPEN_BORDER))
          
60        Rectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight

          Dim textRect As RECT
          
70        textRect.left = 3
80        textRect.top = 3
90        textRect.right = UserControl.ScaleWidth - 3
100       textRect.bottom = UserControl.ScaleHeight - 3
          
          Dim focusRect As RECT
          
110       focusRect.left = 2
120       focusRect.top = 2
130       focusRect.right = UserControl.ScaleWidth - 2
140       focusRect.bottom = UserControl.ScaleHeight - 2
          
150       If m_hasFocus Then
160           SetTextColor UserControl.hdc, Not colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
170           DrawFocusRect UserControl.hdc, focusRect
180       End If
          
190       If m_state = bsNormal Then
200           SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
210       ElseIf m_state = bsMouseover Then
220           SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
230       ElseIf m_state = bsMouseDown Then
240           textRect.left = 7
250           textRect.top = 7
260           SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
270       End If
          
280       swiftDrawText UserControl.hdc, m_caption, VarPtr(textRect), DT_SINGLELINE Or DT_VCENTER Or _
              DT_CENTER Or DT_END_ELLIPSIS
              
290       SelectObject UserControl.hdc, oldBrush
300       SelectObject UserControl.hdc, oldPen
End Sub
