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
    isUrl = m_isUrl
End Property

Public Property Let isUrl(newValue As Boolean)
    m_isUrl = newValue
End Property

Public Property Get url() As String
    url = m_url
End Property

Public Property Let url(newValue As String)
    m_url = newValue
End Property

Public Property Get text() As String
    text = m_text
End Property

Public Property Let text(newValue As String)
    m_text = newValue
End Property

Public Property Get left() As Long
    left = m_left
End Property

Public Property Let left(newValue As Long)
    m_left = newValue
End Property

Public Property Get width() As Long
    width = m_width
End Property

Public Property Let width(newValue As Long)
    m_width = newValue
End Property

Private Function ITextRenderBlock_getSelectedText(drawingData As CDrawingData, withFormatCodes As _
    Boolean) As String
    
    Dim selectedAmount As Long
    
    If m_selected Then
        If m_allSelected Or (m_selectLeft < 1 And m_selectRight >= m_width) Then
            ITextRenderBlock_getSelectedText = m_text
            Exit Function
        End If
        
        Dim maxFit As Long
        Dim textSize As SIZE
        Dim selStart As Long
        Dim selEnd As Long
        
        If m_selectLeft > 0 Then
            swiftGetTextExtentExPoint drawingData.Dc, m_text, m_selectLeft, maxFit, textSize
            
            selStart = maxFit + 1
            
            If m_selectRight < m_width Then
                swiftGetTextExtentExPoint drawingData.Dc, StrReverse(m_text), m_width - _
                    m_selectRight, maxFit, textSize
                    
                selEnd = (Len(m_text) - (maxFit) + 1)
                
                selectedAmount = selEnd - selStart
                
                If selectedAmount < 1 Then
                    Exit Function
                End If
                
                ITextRenderBlock_getSelectedText = Mid$(m_text, selStart, selectedAmount)
            Else
                ITextRenderBlock_getSelectedText = Mid$(m_text, selStart)
            End If
        Else
            swiftGetTextExtentExPoint drawingData.Dc, m_text, m_selectRight, maxFit, textSize
                
            ITextRenderBlock_getSelectedText = Mid$(m_text, 1, maxFit + 1)
        End If
    End If
End Function

Private Function ITextRenderBlock_getText() As String
    ITextRenderBlock_getText = m_text
End Function

Private Sub ITextRenderBlock_render(drawingData As CDrawingData)
    Dim oldColour As Long
    Dim oldHasColour As Boolean
    Dim oldUnderline As Boolean
    Dim ret As Long
    
    If m_isUrl Then
        oldHasColour = drawingData.hasForeColour
        oldColour = drawingData.foreColour
        oldUnderline = drawingData.underline
        
        drawingData.foreColour = eventColours.urls.colour
        drawingData.hasForeColour = True
        drawingData.underline = True
    End If
    
    Dim textRect As RECT
    
    textRect.left = m_left
    textRect.right = m_left + m_width
    textRect.top = drawingData.y
    textRect.bottom = drawingData.y + drawingData.fontHeight
    
    drawingData.x = m_left + m_width
    
    If m_selected Then
        If m_allSelected Or (m_selectLeft < 1 And m_selectRight >= m_width) Then
            drawingData.reverse = Not drawingData.reverse
            swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, VarPtr(textRect), _
                m_text
            drawingData.reverse = Not drawingData.reverse
        Else
            Dim maxFit As Long
            Dim drawnWidth As Long
            Dim textIndex As Long
            Dim textSize As SIZE
            Dim text As String
            Dim selectStart As Long
            Dim selectEnd As Long
            
            If m_selectLeft > 0 Then
                ret = swiftGetTextExtentExPoint(drawingData.Dc, m_text, m_selectLeft, maxFit, _
                    textSize)
                selectStart = maxFit + 1
                
                If selectStart > 1 Then
                    text = Mid$(m_text, 1, maxFit)
                    swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, _
                        VarPtr(textRect), text
                    
                    swiftGetTextExtentPoint32 drawingData.Dc, text, textSize
                    drawnWidth = textSize.cx
                End If
                
                If m_selectRight < m_width Then
                    swiftGetTextExtentExPoint drawingData.Dc, StrReverse(m_text), m_width - _
                        m_selectRight, maxFit, textSize
            
                    selectEnd = Len(m_text) - (maxFit - 1)
                    
                    If selectEnd > selectStart Then
                        text = Mid$(m_text, selectStart, selectEnd - selectStart)
                    Else
                        text = vbNullString
                    End If
                Else
                    text = Mid$(m_text, selectStart)
                End If
                
                If LenB(text) <> 0 Then
                    drawingData.reverse = Not drawingData.reverse
                    swiftTextOut drawingData.Dc, m_left + drawnWidth, drawingData.y, ETO_CLIPPED, _
                        VarPtr(textRect), text
                    drawingData.reverse = Not drawingData.reverse
                    
                    swiftGetTextExtentPoint32 drawingData.Dc, text, textSize
                    drawnWidth = drawnWidth + textSize.cx
                End If
                
                If selectEnd Then
                    text = Mid$(m_text, selectEnd)
                    swiftTextOut drawingData.Dc, m_left + drawnWidth, drawingData.y, ETO_CLIPPED, 0, _
                        text
                End If
            Else
                drawingData.reverse = Not drawingData.reverse
                swiftGetTextExtentExPoint drawingData.Dc, m_text, m_selectRight, maxFit, textSize
                text = Mid$(m_text, 1, maxFit + 1)
                swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, VarPtr(textRect), _
                    text
                drawingData.reverse = Not drawingData.reverse
                
                swiftGetTextExtentPoint32 drawingData.Dc, text, textSize
                text = Mid$(m_text, maxFit + 2)
                swiftTextOut drawingData.Dc, m_left + textSize.cx, drawingData.y, ETO_CLIPPED, _
                    VarPtr(textRect), text
            End If
        End If
    Else
        swiftTextOut drawingData.Dc, m_left, drawingData.y, ETO_CLIPPED, VarPtr(textRect), m_text
    End If
    
    If m_isUrl Then
        drawingData.foreColour = oldColour
        drawingData.hasForeColour = oldHasColour
        drawingData.underline = oldUnderline
    End If
End Sub

Private Sub ITextRenderBlock_selectAll()
    m_selected = True
    m_allSelected = True
End Sub

Private Property Get ITextRenderBlock_selected() As Boolean
    ITextRenderBlock_selected = m_selected
End Property

Private Sub ITextRenderBlock_selectSome(left As Long, right As Long)
    m_selected = True
    m_allSelected = False
    
    m_selectLeft = left - m_left
    
    If right = -1 Then
        m_selectRight = m_width
    ElseIf right < m_left Then
        m_selected = False
    Else
        m_selectRight = right - m_left
    End If
End Sub

Private Sub ITextRenderBlock_unSelect()
    m_selected = False
    m_allSelected = False
    m_selectLeft = 0
    m_selectRight = 0
End Sub

Private Property Get ITextRenderBlock_width() As Long
    ITextRenderBlock_width = m_width
End Property

Private Property Let ITextRenderBlock_width(RHS As Long)
    m_width = RHS
End Property

Private Property Get ITextRenderBlock_left() As Long
    ITextRenderBlock_left = m_left
End Property

Private Property Let ITextRenderBlock_left(RHS As Long)
    m_left = RHS
End Property
