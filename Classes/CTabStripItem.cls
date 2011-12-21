VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTabStripItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_caption As String
Private m_prefix As String
Private m_state As eTabStripItemState

Public Sub init(caption As String, prefix As String)
    m_caption = caption
    m_prefix = prefix
End Sub

Public Property Get caption() As String
    caption = m_caption
End Property

Public Property Get prefix() As String
    prefix = m_prefix
End Property

Public Property Get state() As eTabStripItemState
    state = m_state
End Property

Public Property Let state(newValue As eTabStripItemState)
    m_state = newValue
End Property

Public Sub render(hdc As Long, x As Long, y As Long, width As Long, height As Long)
    Dim oldBrush As Long
    Dim oldPen As Long
    
    If m_state = tisSelected Then
        oldBrush = SelectObject(hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK))
        oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_FRAMEBACK))
    Else
        oldBrush = SelectObject(hdc, colourManager.getBrush(SWIFTCOLOUR_WINDOW))
        oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_FRAMEBORDER))
    End If
    
    RoundRect hdc, x, y, x + width, y + height, 10, 10
    
    If m_state = tisMouseOver Or m_state = tisSelected Then
        SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
    Else
        SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
    End If
    
    SetBkMode hdc, TRANSPARENT
    
    swiftDrawText hdc, m_caption, VarPtr(makeRect(x + 2, (x - 2) + width, y, y + height)), _
        DT_SINGLELINE Or DT_VCENTER Or DT_CENTER Or DT_END_ELLIPSIS
    
    SelectObject hdc, oldBrush
    SelectObject hdc, oldPen
End Sub
