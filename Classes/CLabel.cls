VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_caption As String
Private m_x As Long
Private m_y As Long
Private m_labelType As eLabelType
Private m_visible As Boolean
Private m_font As Long
Private m_gotRect As Boolean
Private m_rect As RECT

Public Property Get font() As Long
10        font = m_font
End Property

Public Property Let font(newValue As Long)
10        m_font = newValue
End Property

Public Property Get visible() As Boolean
10        visible = m_visible
End Property

Public Property Let visible(newValue As Boolean)
10        m_visible = newValue
End Property

Public Property Get caption() As String
10        caption = m_caption
End Property

Public Property Let caption(newValue As String)
10        m_caption = newValue
End Property

Friend Sub setRect(newRect As RECT)
10        m_gotRect = True
20        m_rect = newRect
End Sub

Friend Sub init(caption As String, labelType As eLabelType, x As Long, y As Long)
10        m_caption = caption
20        m_labelType = labelType
30        m_x = x
40        m_y = y
50        m_visible = True
End Sub

Friend Sub render(hdc As Long)
          Dim oldFont As Long
          
10        oldFont = GetCurrentObject(hdc, OBJ_FONT)
          
20        If m_font <> 0 Then
30            SelectObject hdc, m_font
40        ElseIf m_labelType = ltHeading Then
50            SelectObject hdc, g_fontHeading
60        ElseIf m_labelType = ltSubHeading Then
70            SelectObject hdc, g_fontSubHeading
80        Else
90            SelectObject hdc, g_fontUI
100       End If
          
110       SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
          
          Dim textSize As SIZE
          
120       If Not m_gotRect Then
130           swiftGetTextExtentPoint32 hdc, m_caption, textSize
              
              Dim textRect As RECT
              
140           textRect = makeRect(m_x, m_x + textSize.cx, m_y, m_y + textSize.cy)
150           swiftDrawText hdc, m_caption, VarPtr(textRect), DT_SINGLELINE Or DT_VCENTER
160       Else
170           swiftDrawText hdc, m_caption, VarPtr(m_rect), DT_WORDBREAK
180       End If
          
190       SelectObject hdc, oldFont
End Sub
