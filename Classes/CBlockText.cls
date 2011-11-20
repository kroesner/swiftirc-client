VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBlockText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextRenderBlock

Private m_width As Long
Private m_left As Long

Private m_selected As Boolean
Private m_allSelected As Boolean

Private m_selectLeft As Long
Private m_selectRight As Long

Private m_isUrl As Boolean
Private m_url As String

Private m_text As String

Public Property Get isUrl() As Boolean
10        isUrl = m_isUrl
End Property

Public Property Let isUrl(newValue As Boolean)
10        m_isUrl = newValue
End Property

Public Property Get url() As String
10        url = m_url
End Property

Public Property Let url(newValue As String)
10        m_url = newValue
End Property

Public Property Get text() As String
10        text = m_text
End Property

Public Property Let text(newValue As String)
10        m_text = newValue
End Property

Public Property Get left() As Long
10        left = m_left
End Property

Public Property Let left(newValue As Long)
10        m_left = newValue
End Property

Public Property Get width() As Long
10        width = m_width
End Property

Public Property Let width(newValue As Long)
10        m_width = newValue
End Property

Private Function ITextRenderBlock_getSelectedText(drawingData As CDrawingData, withFormatCodes As _
    Boolean) As String
          
          Dim selectedAmount As Long
          
10        If m_selected Then
20            If m_allSelected Or (m_selectLeft < 1 And m_selectRight >= m_width) Then
30                ITextRenderBlock_getSelectedText = m_text
40                Exit Function
50            End If
              
              Dim maxFit As Long
              Dim textSize As SIZE
              Dim selStart As Long
              Dim selEnd As Long
              
60            If m_selectLeft > 0 Then
70                swiftGetTextExtentExPoint drawingData.Dc, m_text, m_selectLeft, maxFit, textSize
                  
80                selStart = maxFit + 1
                  
90                If m_selectRight < m_width Then
100                   swiftGetTextExtentExPoint drawingData.Dc, StrReverse(m_text), m_width - _
                          m_selectRight, maxFit, textSize
                          
110                   selEnd = (Len(m_text) - (maxFit) + 1)
                      
120                   selectedAmount = selEnd - selStart
                      
130                   If selectedAmount < 1 Then
140                       Exit Function
150                   End If
                      
160                   ITextRenderBlock_getSelectedText = Mid$(m_text, selStart, selectedAmount)
170               Else
180                   ITextRenderBlock_getSelectedText = Mid$(m_text, selStart)
190               End If
200           Else
210               swiftGetTextExtentExPoint drawingData.Dc, m_text, m_selectRight, maxFit, textSize
                      
220               ITextRenderBlock_getSelectedText = Mid$(m_text, 1, maxFit + 1)
230           End If
240       End If
End Function

Private Function ITextRenderBlock_getText() As String
10        ITextRenderBlock_getText = m_text
End Function

Private Sub ITextRenderBlock_render(drawingData As CDrawingData)
          Dim oldColour As Long
          Dim oldHasColour As Boolean
          Dim oldUnderline As Boolean
          Dim ret As Long
          
10        If m_isUrl Then
20            oldHasColour = drawingData.hasForeColour
30            oldColour = drawingData.foreColour
40            oldUnderline = drawingData.underline
              
50            drawingData.foreColour = eventColours.urls.colour
60            drawingData.hasForeColour = True
70            drawingData.underline = True
80        End If
          
          Dim textRect As RECT
          
90        textRect.left = m_left
100       textRect.right = m_left + m_width
110       textRect.top = drawingData.y
120       textRect.bottom = drawingData.y + drawingData.fontHeight
          
130       drawingData.x = m_left + m_width
          
140       If m_selected Then
150           If m_allSelected Or (m_selectLeft < 1 And m_selectRight >= m_width) Then
160               drawingData.reverse = Not drawingData.reverse
170               swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, VarPtr(textRect), _
                      m_text
180               drawingData.reverse = Not drawingData.reverse
190           Else
                  Dim maxFit As Long
                  Dim drawnWidth As Long
                  Dim textIndex As Long
                  Dim textSize As SIZE
                  Dim text As String
                  Dim selectStart As Long
                  Dim selectEnd As Long
                  
200               If m_selectLeft > 0 Then
210                   ret = swiftGetTextExtentExPoint(drawingData.Dc, m_text, m_selectLeft, maxFit, _
                          textSize)
220                   selectStart = maxFit + 1
                      
230                   If selectStart > 1 Then
240                       text = Mid$(m_text, 1, maxFit)
250                       swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, _
                              VarPtr(textRect), text
                          
260                       swiftGetTextExtentPoint32 drawingData.Dc, text, textSize
270                       drawnWidth = textSize.cx
280                   End If
                      
290                   If m_selectRight < m_width Then
300                       swiftGetTextExtentExPoint drawingData.Dc, StrReverse(m_text), m_width - _
                              m_selectRight, maxFit, textSize
                  
310                       selectEnd = Len(m_text) - (maxFit - 1)
                          
320                       If selectEnd > selectStart Then
330                           text = Mid$(m_text, selectStart, selectEnd - selectStart)
340                       Else
350                           text = vbNullString
360                       End If
370                   Else
380                       text = Mid$(m_text, selectStart)
390                   End If
                      
400                   If LenB(text) <> 0 Then
410                       drawingData.reverse = Not drawingData.reverse
420                       swiftTextOut drawingData.Dc, m_left + drawnWidth, drawingData.y, ETO_CLIPPED, _
                              VarPtr(textRect), text
430                       drawingData.reverse = Not drawingData.reverse
                          
440                       swiftGetTextExtentPoint32 drawingData.Dc, text, textSize
450                       drawnWidth = drawnWidth + textSize.cx
460                   End If
                      
470                   If selectEnd Then
480                       text = Mid$(m_text, selectEnd)
490                       swiftTextOut drawingData.Dc, m_left + drawnWidth, drawingData.y, ETO_CLIPPED, 0, _
                              text
500                   End If
510               Else
520                   drawingData.reverse = Not drawingData.reverse
530                   swiftGetTextExtentExPoint drawingData.Dc, m_text, m_selectRight, maxFit, textSize
540                   text = Mid$(m_text, 1, maxFit + 1)
550                   swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, VarPtr(textRect), _
                          text
560                   drawingData.reverse = Not drawingData.reverse
                      
570                   swiftGetTextExtentPoint32 drawingData.Dc, text, textSize
580                   text = Mid$(m_text, maxFit + 2)
590                   swiftTextOut drawingData.Dc, m_left + textSize.cx, drawingData.y, ETO_CLIPPED, _
                          VarPtr(textRect), text
600               End If
610           End If
620       Else
630           swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, VarPtr(textRect), m_text
640       End If
          
650       If m_isUrl Then
660           drawingData.foreColour = oldColour
670           drawingData.hasForeColour = oldHasColour
680           drawingData.underline = oldUnderline
690       End If
End Sub

Private Sub ITextRenderBlock_selectAll()
10        m_selected = True
20        m_allSelected = True
End Sub

Private Property Get ITextRenderBlock_selected() As Boolean
10        ITextRenderBlock_selected = m_selected
End Property

Private Sub ITextRenderBlock_selectSome(left As Long, right As Long)
10        m_selected = True
20        m_allSelected = False
          
30        m_selectLeft = left - m_left
          
40        If right = -1 Then
50            m_selectRight = m_width
60        ElseIf right < m_left Then
70            m_selected = False
80        Else
90            m_selectRight = right - m_left
100       End If
End Sub

Private Sub ITextRenderBlock_unSelect()
10        m_selected = False
20        m_allSelected = False
30        m_selectLeft = 0
40        m_selectRight = 0
End Sub

Private Property Get ITextRenderBlock_width() As Long
10        ITextRenderBlock_width = m_width
End Property

Private Property Let ITextRenderBlock_width(RHS As Long)
10        m_width = RHS
End Property

Private Property Get ITextRenderBlock_left() As Long
10        ITextRenderBlock_left = m_left
End Property

Private Property Let ITextRenderBlock_left(RHS As Long)
10        m_left = RHS
End Property
