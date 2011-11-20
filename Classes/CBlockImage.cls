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
10        caption = m_caption
End Property

Public Property Let caption(newValue As String)
10        m_caption = newValue
End Property

Private Function ITextRenderBlock_getSelectedText(drawingData As CDrawingData, withFormatCodes As _
    Boolean) As String
          
10        If m_selected Then
20            ITextRenderBlock_getSelectedText = m_caption
30        End If
End Function

Private Property Let ITextRenderBlock_left(RHS As Long)
10        m_left = RHS
End Property

Private Property Get ITextRenderBlock_left() As Long
10        ITextRenderBlock_left = m_left
End Property

Public Property Let image(newImage As CImage)
10        Set m_image = newImage
End Property

Private Sub ITextRenderBlock_render(drawingData As CDrawingData)
10        If m_image Is Nothing Then
20            Exit Sub
30        End If
          
          Dim tempDc As Long
          Dim tempBitmap As Long
          Dim oldBitmap As Long
          
40        tempDc = CreateCompatibleDC(drawingData.Dc)
50        tempBitmap = CreateCompatibleBitmap(drawingData.Dc, drawingData.fontHeight, drawingData.fontHeight)
          
60        oldBitmap = SelectObject(tempDc, tempBitmap)
          
          Dim imageRect As RECT
          
70        imageRect = makeRect(0, drawingData.fontHeight, 0, drawingData.fontHeight)
          
80        If m_selected Then
90            FillRect tempDc, imageRect, drawingData.altFillBrush
100       Else
110           FillRect tempDc, imageRect, drawingData.fillBrush
120       End If
          
130       If m_selected Then
140           m_image.drawInverted tempDc, 0, 0, drawingData.fontHeight, drawingData.fontHeight
150       Else
160           m_image.draw tempDc, 0, 0, drawingData.fontHeight, drawingData.fontHeight
170       End If
          
180       BitBlt drawingData.Dc, m_left, drawingData.y, drawingData.fontHeight, drawingData.fontHeight, _
                  tempDc, 0, 0, vbSrcCopy
          
190       SelectObject tempDc, oldBitmap
          
200       DeleteDC tempDc
210       DeleteObject tempBitmap
          
220       drawingData.x = drawingData.x + drawingData.fontHeight
End Sub

Private Sub ITextRenderBlock_selectAll()
10        m_selected = True
End Sub

Private Property Get ITextRenderBlock_selected() As Boolean
10        ITextRenderBlock_selected = m_selected
End Property

Private Sub ITextRenderBlock_selectSome(left As Long, width As Long)
10        m_selected = True
End Sub

Private Sub ITextRenderBlock_unSelect()
10        m_selected = False
End Sub

Private Property Let ITextRenderBlock_width(RHS As Long)
10        m_width = RHS
End Property

Private Property Get ITextRenderBlock_width() As Long
10        ITextRenderBlock_width = m_width
End Property
