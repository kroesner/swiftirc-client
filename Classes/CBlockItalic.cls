VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBlockItalic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextRenderBlock

Private m_italic As Boolean

Private Function ITextRenderBlock_getSelectedText(drawingData As CDrawingData, withFormatCodes As _
    Boolean) As String

End Function

Public Property Get italic() As Boolean
    italic = m_italic
End Property

Public Property Let italic(newValue As Boolean)
    m_italic = newValue
End Property

Private Sub ITextRenderBlock_render(drawingData As CDrawingData)
    drawingData.italic = m_italic
End Sub

Private Property Get ITextRenderBlock_width() As Long
    ITextRenderBlock_width = 0
End Property

Private Property Let ITextRenderBlock_width(RHS As Long)

End Property

Private Property Get ITextRenderBlock_left() As Long

End Property

Private Property Let ITextRenderBlock_left(RHS As Long)

End Property

Private Sub ITextRenderBlock_selectAll()

End Sub

Private Property Get ITextRenderBlock_selected() As Boolean

End Property

Private Sub ITextRenderBlock_selectSome(left As Long, width As Long)

End Sub

Private Sub ITextRenderBlock_unSelect()

End Sub
