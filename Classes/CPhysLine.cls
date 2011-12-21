VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_top As Integer
Private m_bottom As Integer

Private m_selected As Boolean

Private m_blocks As New cArrayList

Public Property Get selected() As Boolean
    selected = m_selected
End Property

Public Sub addBlock(ByRef block As ITextRenderBlock)
    m_blocks.Add block
End Sub

Public Sub cloneFormat(drawingData As CDrawingData)
    If drawingData.bold Then
        Dim blockBold As New CBlockBold
        
        blockBold.bold = True
        m_blocks.Add blockBold
    End If
    
    If drawingData.italic Then
        Dim blockItalic As New CBlockItalic
        
        blockItalic.italic = True
        m_blocks.Add blockItalic
    End If
    
    If drawingData.underline Then
        Dim blockUnderline As New CBlockUnderline
        
        blockUnderline.underline = True
        m_blocks.Add blockUnderline
    End If
    
    If drawingData.reverse Then
        Dim blockReverse As New CBlockReverse
        
        blockReverse.reverse = True
        m_blocks.Add blockReverse
    End If
    
    If drawingData.hasForeColour Then
        Dim blockForeColour As New CBlockForeColour
        
        blockForeColour.hasForeColour = True
        blockForeColour.foreColour = drawingData.foreColour
        m_blocks.Add blockForeColour
    End If
    
    If drawingData.hasBackColour Then
        Dim blockbackColour As New CBlockBackColour
        
        blockbackColour.hasBackColour = True
        blockbackColour.backColour = drawingData.backColour
        m_blocks.Add blockbackColour
    End If
End Sub

Public Function block(index As Integer) As ITextRenderBlock
    Set block = m_blocks.item(index)
End Function

Public Function blockCount() As Integer
    blockCount = m_blocks.count
End Function

Public Property Get top() As Integer
    top = m_top
End Property

Public Property Let top(newValue As Integer)
    m_top = newValue
End Property

Public Property Get bottom() As Integer
    bottom = m_bottom
End Property

Public Property Let bottom(newValue As Integer)
    m_bottom = newValue
End Property

Public Sub setSelection(ByVal left As Long, ByVal right As Long)
    Dim count As Integer
    Dim block As ITextRenderBlock
    Dim selectStart As Boolean
    Dim selectEnd As Boolean
    
    For count = 1 To m_blocks.count
        Set block = m_blocks.item(count)
        
        If block.width <> 0 Then
            If selectEnd Then
                If block.selected Then block.unSelect
            ElseIf selectStart Then
                If right <= block.left + block.width Then
                    block.selectSome left, right
                    selectEnd = True
                Else
                    block.selectAll
                End If
            Else
                If block.left + block.width >= left Then
                    selectStart = True
                    block.selectSome left, right
                
                    If right = -1 Then
                        Dim count2 As Integer
                    
                        For count2 = count + 1 To m_blocks.count
                            Set block = m_blocks.item(count2)
                            block.selectAll
                        Next count2
                        
                        Exit For
                    End If
                
                    If right <= block.left + block.width Then
                        selectEnd = True
                    End If
                Else
                    If block.selected Then block.unSelect
                End If
            End If
        End If
    Next count
    
    m_selected = True
End Sub

Public Sub selectAll()
    Dim count As Integer
    Dim block As ITextRenderBlock
    
    For count = 1 To m_blocks.count
        Set block = m_blocks.item(count)
        block.selectAll
    Next count
    
    m_selected = True
End Sub

Public Sub unSelect()
    Dim count As Integer
    Dim block As ITextRenderBlock

    For count = 1 To m_blocks.count
        Set block = m_blocks.item(count)
        block.unSelect
    Next count
    
    m_selected = False
End Sub

Public Function getSelectedText(drawingData As CDrawingData, withFormatCodes As Boolean) As String
    Dim count As Integer
    Dim block As ITextRenderBlock
    
    Dim text As String
    
    For count = 1 To m_blocks.count
        Set block = m_blocks.item(count)
        text = text & block.getSelectedText(drawingData, withFormatCodes)
    Next count
    
    getSelectedText = text
End Function

Friend Function getMouseOverBlock(ByVal x As Long) As ITextRenderBlock
    Dim count As Long
    Dim block As ITextRenderBlock
    
    For count = 1 To m_blocks.count
        Set block = m_blocks.item(count)
        
        If block.width <> 0 Then
            If x >= m_blocks.item(count).left Then
                If x <= m_blocks.item(count).left + m_blocks.item(count).width Then
                    Set getMouseOverBlock = block
                    Exit Function
                End If
            End If
        End If
    Next count
End Function

Public Sub render(drawingData As CDrawingData)
    Dim count As Integer
    Dim block As ITextRenderBlock
    
    For count = 1 To m_blocks.count
        Set block = m_blocks.item(count)
        
        If block.width <> 0 Then
            Exit For
        End If
    Next count
    
    If block Is Nothing Then
        Exit Sub
    End If
    
    If block.left > 0 Then
       'Indented
        drawingData.fillLineIndent block.left
    End If
    
    For count = 1 To m_blocks.count
        Set block = m_blocks.item(count)
        block.render drawingData
    Next count
    
    m_top = drawingData.realY - drawingData.fontHeight
    m_bottom = drawingData.realY
    
    drawingData.fillRemainingLine
End Sub
