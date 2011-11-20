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
10        enabled = m_realWindow.enabled
End Property

Public Property Let enabled(newEnabled As Boolean)
10        m_realWindow.enabled = False
End Property

Public Property Get password() As Boolean
10        password = m_password
End Property

Public Property Let password(newValue As Boolean)
10        m_password = newValue
          
20        If m_password Then
30            m_textBox.PasswordChar = "*"
40        Else
50            m_textBox.PasswordChar = vbNullString
60        End If
End Property

Public Property Get visible() As Boolean
10        visible = m_realWindow.visible
End Property

Public Property Let visible(newValue As Boolean)
10        m_realWindow.visible = newValue
End Property

Public Property Get value() As String
10        value = m_textBox.text
End Property

Public Property Let value(newValue As String)
10        m_textBox.text = newValue
End Property

Public Property Get required() As Boolean
10        required = m_required
End Property

Public Property Let required(newValue As Boolean)
10        m_required = newValue
End Property

Public Property Get justification() As eFieldJustification
10        justification = m_fieldJustification
End Property

Public Property Let justification(newValue As eFieldJustification)
10        m_fieldJustification = newValue
End Property

Public Sub setFieldWidth(captionWidth As Integer, boxWidth As Integer)
10        m_captionWidth = captionWidth
20        m_BoxWidth = boxWidth
30        UserControl_Resize
End Sub

Public Property Get caption() As String
10        caption = m_caption
End Property

Public Property Let caption(newValue As String)
10        m_caption = newValue
20        UserControl_Paint
End Property

Public Property Get mask() As eFieldMask
10        mask = m_fieldMask
End Property

Public Property Let mask(newValue As eFieldMask)
10        m_fieldMask = newValue
End Property

Private Sub IColourUser_coloursUpdated()
10        updateColours
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub m_textbox_KeyPress(KeyAscii As Integer)
10        If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
20            Exit Sub
30        End If

40        If m_fieldMask = fmNumericOnly Then
50            If KeyAscii < 48 Or KeyAscii > 57 Then
60                KeyAscii = 0
70            End If
80        ElseIf m_fieldMask = fmAlphaOnly Then
90            If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) Then
100               KeyAscii = 0
110           End If
120       ElseIf m_fieldMask = fmIrcNickname Then
130           If KeyAscii = vbKeySpace Then
140               KeyAscii = Asc("_")
150           ElseIf (KeyAscii > 47 And KeyAscii < 58) Then
160               If LenB(m_textBox.text) = 0 Then
170                   KeyAscii = 0
180               End If
190           ElseIf KeyAscii = Asc("@") Then
200               KeyAscii = 0
210           ElseIf KeyAscii = 34 Then
220               KeyAscii = 0
230           ElseIf KeyAscii = Asc("/") Then
240               KeyAscii = 0
250           ElseIf KeyAscii = Asc("*") Then
260               KeyAscii = 0
270           End If
280       End If
End Sub

Private Sub UserControl_Initialize()
10        SetBkMode UserControl.hdc, TRANSPARENT
20        Set m_textBox = Controls.Add("VB.TextBox", "textbox")
30        m_textBox.Appearance = 0
40        m_textBox.BorderStyle = 0
50        m_textBox.visible = True
          
60        SendMessage m_textBox.hwnd, WM_SETFONT, g_fontUI, 1
70        SelectObject UserControl.hdc, g_fontSubHeading
          
80        updateColours
End Sub

Private Sub UserControl_Paint()
          Dim labelSize As Long
          
10        If m_captionWidth <> 0 Then
20            labelSize = m_captionWidth
30            FrameRect UserControl.hdc, makeRect(labelSize, labelSize + m_BoxWidth, 0, _
                  UserControl.ScaleHeight), colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
40        Else
50            labelSize = (UserControl.ScaleWidth / 2)
60            FrameRect UserControl.hdc, makeRect(labelSize, UserControl.ScaleWidth, 0, _
                  UserControl.ScaleHeight), colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
70        End If
          
          Dim textRect As RECT
          
80        textRect = makeRect(0, labelSize - 5, 0, UserControl.ScaleHeight)
          
90        If m_required Then
100           swiftDrawText UserControl.hdc, "*", VarPtr(textRect), DT_SINGLELINE Or DT_VCENTER
110           textRect.left = textRect.left + UserControl.TextWidth("*")
120       End If
          
          Dim justifyFlag As Long
          
130       If m_fieldJustification = fjRight Then
140           justifyFlag = DT_RIGHT
150       Else
160           justifyFlag = 0
170       End If
          
180       swiftDrawText UserControl.hdc, m_caption, VarPtr(textRect), DT_VCENTER Or DT_SINGLELINE Or _
              justifyFlag Or DT_END_ELLIPSIS
End Sub

Private Sub UserControl_Resize()
10        If m_captionWidth <> 0 And m_BoxWidth <> 0 Then
20            m_textBox.Move m_captionWidth + 1, 1, m_BoxWidth - 2, UserControl.ScaleHeight - 2
30        Else
40            m_textBox.Move (UserControl.ScaleWidth / 2) + 1, 1, (UserControl.ScaleWidth / 2) - 2, _
                  UserControl.ScaleHeight - 2
50        End If
End Sub

Private Sub updateColours()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
20        SetTextColor UserControl.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
30        m_textBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
40        m_textBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
End Sub
