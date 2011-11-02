VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBlockImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextRenderBlock

Private m_left As Long
Private m_width As Long
Private m_selected As Boolean
Private m_image As CImage
Private m_caption As String

Public Property Get caption() As String
    caption = m_caption
End Property

Public Property Let caption(newValue As String)
    m_caption = newValue
End Property

Private Function ITextRenderBlock_getSelectedText(drawingData As CDrawingData, withFormatCodes As _
    Boolean) As String
    
    If m_selected Then
        ITextRenderBlock_getSelectedText = m_caption
    End If
End Function

Private Property Let ITextRenderBlock_left(RHS As Long)
    m_left = RHS
End Property

Private Property Get ITextRenderBlock_left() As Long
    ITextRenderBlock_left = m_left
End Property

Public Property Let image(newImage As CImage)
    Set m_image = newImage
End Property

Private Sub ITextRenderBlock_render(drawingData As CDrawingData)
    If m_image Is Nothing Then
        Exit Sub
    End If
    
    Dim tempDc As Long
    Dim tempBitmap As Long
    Dim oldBitmap As Long
    
    tempDc = CreateCompatibleDC(drawingData.Dc)
    tempBitmap = CreateCompatibleBitmap(drawingData.Dc, drawingData.fontHeight, drawingData.fontHeight)
    
    oldBitmap = SelectObject(tempDc, tempBitmap)
    
    Dim imageRect As RECT
    
    imageRect = makeRect(0, drawingData.fontHeight, 0, drawingData.fontHeight)
    
    If m_selected Then
        FillRect tempDc, imageRect, drawingData.altFillBrush
    Else
        FillRect tempDc, imageRect, drawingData.fillBrush
    End If
    
    If m_selected Then
        m_image.drawInverted tempDc, 0, 0, drawingData.fontHeight, drawingData.fontHeight
    Else
        m_image.draw tempDc, 0, 0, drawingData.fontHeight, drawingData.fontHeight
    End If
    
    BitBlt drawingData.Dc, m_left, drawingData.y, drawingData.fontHeight, drawingData.fontHeight, _
            tempDc, 0, 0, vbSrcCopy
    
    SelectObject tempDc, oldBitmap
    
    DeleteDC tempDc
    DeleteObject tempBitmap
    
    drawingData.x = drawingData.x + drawingData.fontHeight
End Sub

Private Sub ITextRenderBlock_selectAll()
    m_selected = True
End Sub

Private Property Get ITextRenderBlock_selected() As Boolean
    ITextRenderBlock_selected = m_selected
End Property

Private Sub ITextRenderBlock_selectSome(left As Long, width As Long)
    m_selected = True
End Sub

Private Sub ITextRenderBlock_unSelect()
    m_selected = False
End Sub

Private Property Let ITextRenderBlock_width(RHS As Long)
    m_width = RHS
End Property

Private Property Get ITextRenderBlock_width() As Long
    ITextRenderBlock_width = m_width
End Property
