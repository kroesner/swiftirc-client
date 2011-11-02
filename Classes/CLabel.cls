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
    font = m_font
End Property

Public Property Let font(newValue As Long)
    m_font = newValue
End Property

Public Property Get visible() As Boolean
    visible = m_visible
End Property

Public Property Let visible(newValue As Boolean)
    m_visible = newValue
End Property

Public Property Get caption() As String
    caption = m_caption
End Property

Public Property Let caption(newValue As String)
    m_caption = newValue
End Property

Friend Sub setRect(newRect As RECT)
    m_gotRect = True
    m_rect = newRect
End Sub

Friend Sub init(caption As String, labelType As eLabelType, x As Long, y As Long)
    m_caption = caption
    m_labelType = labelType
    m_x = x
    m_y = y
    m_visible = True
End Sub

Friend Sub render(hdc As Long)
    Dim oldFont As Long
    
    oldFont = GetCurrentObject(hdc, OBJ_FONT)
    
    If m_font <> 0 Then
        SelectObject hdc, m_font
    ElseIf m_labelType = ltHeading Then
        SelectObject hdc, g_fontHeading
    ElseIf m_labelType = ltSubHeading Then
        SelectObject hdc, g_fontSubHeading
    Else
        SelectObject hdc, g_fontUI
    End If
    
    SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
    
    Dim textSize As SIZE
    
    If Not m_gotRect Then
        swiftGetTextExtentPoint32 hdc, m_caption, textSize
        
        Dim textRect As RECT
        
        textRect = makeRect(m_x, m_x + textSize.cx, m_y, m_y + textSize.cy)
        swiftDrawText hdc, m_caption, VarPtr(textRect), DT_SINGLELINE Or DT_VCENTER
    Else
        swiftDrawText hdc, m_caption, VarPtr(m_rect), DT_WORDBREAK
    End If
    
    SelectObject hdc, oldFont
End Sub
