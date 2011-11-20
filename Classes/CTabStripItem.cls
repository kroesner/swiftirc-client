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
10        m_caption = caption
20        m_prefix = prefix
End Sub

Public Property Get caption() As String
10        caption = m_caption
End Property

Public Property Get prefix() As String
10        prefix = m_prefix
End Property

Public Property Get state() As eTabStripItemState
10        state = m_state
End Property

Public Property Let state(newValue As eTabStripItemState)
10        m_state = newValue
End Property

Public Sub render(hdc As Long, x As Long, y As Long, width As Long, height As Long)
          Dim oldBrush As Long
          Dim oldPen As Long
          
10        If m_state = tisSelected Then
20            oldBrush = SelectObject(hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK))
30            oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_FRAMEBACK))
40        Else
50            oldBrush = SelectObject(hdc, colourManager.getBrush(SWIFTCOLOUR_WINDOW))
60            oldPen = SelectObject(hdc, colourManager.getPen(SWIFTPEN_FRAMEBORDER))
70        End If
          
80        RoundRect hdc, x, y, x + width, y + height, 10, 10
          
90        If m_state = tisMouseOver Or m_state = tisSelected Then
100           SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFOREOVER)
110       Else
120           SetTextColor hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
130       End If
          
140       SetBkMode hdc, TRANSPARENT
          
150       swiftDrawText hdc, m_caption, VarPtr(makeRect(x + 2, (x - 2) + width, y, y + height)), _
              DT_SINGLELINE Or DT_VCENTER Or DT_CENTER Or DT_END_ELLIPSIS
          
160       SelectObject hdc, oldBrush
170       SelectObject hdc, oldPen
End Sub
