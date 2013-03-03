VERSION 5.00
Begin VB.UserControl ctlField 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   157
End
Attribute VB_Name = "ctlField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IColourUser
Implements IWindow


Public Enum eFieldMask
    fmAny
    fmNumericOnly
    fmAlphaOnly
    fmIrcNickname
End Enum

Private m_realWindow As VBControlExtender

Private m_caption As String
Private WithEvents m_textBox As VB.TextBox
Attribute m_textBox.VB_VarHelpID = -1
Private m_fieldMask As eFieldMask
Private m_fieldJustification As eFieldJustification
Private m_required As Boolean

Private m_password As Boolean

Private m_captionWidth As Integer
Private m_BoxWidth As Integer

Public Property Get enabled() As Boolean
    enabled = m_realWindow.enabled
End Property

Public Property Let enabled(newEnabled As Boolean)
    m_realWindow.enabled = False
End Property

Public Property Get password() As Boolean
    password = m_password
End Property

Public Property Let password(newValue As Boolean)
    m_password = newValue
    
    If m_password Then
        m_textBox.PasswordChar = "*"
    Else
        m_textBox.PasswordChar = vbNullString
    End If
End Property

Public Property Get visible() As Boolean
    visible = m_realWindow.visible
End Property

Public Property Let visible(newValue As Boolean)
    m_realWindow.visible = newValue
End Property

Public Property Get value() As String
    value = m_textBox.text
End Property

Public Property Let value(newValue As String)
    m_textBox.text = newValue
End Property

Public Property Get required() As Boolean
    required = m_required
End Property

Public Property Let required(newValue As Boolean)
    m_required = newValue
End Property

Public Property Get justification() As eFieldJustification
    justification = m_fieldJustification
End Property

Public Property Let justification(newValue As eFieldJustification)
    m_fieldJustification = newValue
End Property

Public Sub setFieldWidth(captionWidth As Integer, boxWidth As Integer)
    m_captionWidth = captionWidth
    m_BoxWidth = boxWidth
    UserControl_Resize
End Sub

Public Property Get caption() As String
    caption = m_caption
End Property

Public Property Let caption(newValue As String)
    m_caption = newValue
    UserControl_Paint
End Property

Public Property Get mask() As eFieldMask
    mask = m_fieldMask
End Property

Public Property Let mask(newValue As eFieldMask)
    m_fieldMask = newValue
End Property

Private Sub IColourUser_coloursUpdated()
    updateColours
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_textbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    End If

    If m_fieldMask = fmNumericOnly Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    ElseIf m_fieldMask = fmAlphaOnly Then
        If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) Then
            KeyAscii = 0
        End If
    ElseIf m_fieldMask = fmIrcNickname Then
        If KeyAscii = vbKeySpace Then
            KeyAscii = Asc("_")
        ElseIf (KeyAscii > 47 And KeyAscii < 58) Then
            If LenB(m_textBox.text) = 0 Then
                KeyAscii = 0
            End If
        ElseIf KeyAscii = Asc("@") Then
            KeyAscii = 0
        ElseIf KeyAscii = 34 Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc("/") Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc("*") Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    SetBkMode UserControl.hdc, TRANSPARENT
    Set m_textBox = Controls.Add("VB.TextBox", "textbox")
    m_textBox.Appearance = 0
    m_textBox.BorderStyle = 0
    m_textBox.visible = True
    
    SendMessage m_textBox.hwnd, WM_SETFONT, g_fontUI, 1
    SelectObject UserControl.hdc, g_fontSubHeading
    
    updateColours
End Sub

Private Sub UserControl_Paint()
    Dim labelSize As Long
    
    If m_captionWidth <> 0 Then
        labelSize = m_captionWidth
        FrameRect UserControl.hdc, makeRect(labelSize, labelSize + m_BoxWidth, 0, UserControl.ScaleHeight), colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
    Else
        labelSize = (UserControl.ScaleWidth / 2)
        FrameRect UserControl.hdc, makeRect(labelSize, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
    End If
    
    Dim textRect As RECT
    
    textRect = makeRect(0, labelSize - 5, 0, UserControl.ScaleHeight)
    
    If m_required Then
        swiftDrawText UserControl.hdc, "*", VarPtr(textRect), DT_SINGLELINE Or DT_VCENTER
        textRect.left = textRect.left + UserControl.TextWidth("*")
    End If
    
    Dim justifyFlag As Long
    
    If m_fieldJustification = fjRight Then
        justifyFlag = DT_RIGHT
    Else
        justifyFlag = 0
    End If
    
    swiftDrawText UserControl.hdc, m_caption, VarPtr(textRect), DT_VCENTER Or DT_SINGLELINE Or justifyFlag Or DT_END_ELLIPSIS
End Sub

Private Sub UserControl_Resize()
    If m_captionWidth <> 0 And m_BoxWidth <> 0 Then
        m_textBox.Move m_captionWidth + 1, 1, m_BoxWidth - 2, UserControl.ScaleHeight - 2
    Else
        m_textBox.Move (UserControl.ScaleWidth / 2) + 1, 1, (UserControl.ScaleWidth / 2) - 2, UserControl.ScaleHeight - 2
    End If
End Sub

Private Sub updateColours()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
    m_textBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
    m_textBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
End Sub
