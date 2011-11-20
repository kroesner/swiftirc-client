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
10        selected = m_selected
End Property

Public Sub addBlock(ByRef block As ITextRenderBlock)
10        m_blocks.Add block
End Sub

Public Sub cloneFormat(drawingData As CDrawingData)
10        If drawingData.bold Then
              Dim blockBold As New CBlockBold
              
20            blockBold.bold = True
30            m_blocks.Add blockBold
40        End If
          
50        If drawingData.italic Then
              Dim blockItalic As New CBlockItalic
              
60            blockItalic.italic = True
70            m_blocks.Add blockItalic
80        End If
          
90        If drawingData.underline Then
              Dim blockUnderline As New CBlockUnderline
              
100           blockUnderline.underline = True
110           m_blocks.Add blockUnderline
120       End If
          
130       If drawingData.reverse Then
              Dim blockReverse As New CBlockReverse
              
140           blockReverse.reverse = True
150           m_blocks.Add blockReverse
160       End If
          
170       If drawingData.hasForeColour Then
              Dim blockForeColour As New CBlockForeColour
              
180           blockForeColour.hasForeColour = True
190           blockForeColour.foreColour = drawingData.foreColour
200           m_blocks.Add blockForeColour
210       End If
          
220       If drawingData.hasBackColour Then
              Dim blockbackColour As New CBlockBackColour
              
230           blockbackColour.hasBackColour = True
240           blockbackColour.backColour = drawingData.backColour
250           m_blocks.Add blockbackColour
260       End If
End Sub

Public Function block(index As Integer) As ITextRenderBlock
10        Set block = m_blocks.item(index)
End Function

Public Function blockCount() As Integer
10        blockCount = m_blocks.count
End Function

Public Property Get top() As Integer
10        top = m_top
End Property

Public Property Let top(newValue As Integer)
10        m_top = newValue
End Property

Public Property Get bottom() As Integer
10        bottom = m_bottom
End Property

Public Property Let bottom(newValue As Integer)
10        m_bottom = newValue
End Property

Public Sub setSelection(ByVal left As Long, ByVal right As Long)
          Dim count As Integer
          Dim block As ITextRenderBlock
          Dim selectStart As Boolean
          Dim selectEnd As Boolean
          
10        For count = 1 To m_blocks.count
20            Set block = m_blocks.item(count)
              
30            If block.width <> 0 Then
40                If selectEnd Then
50                    If block.selected Then block.unSelect
60                ElseIf selectStart Then
70                    If right <= block.left + block.width Then
80                        block.selectSome left, right
90                        selectEnd = True
100                   Else
110                       block.selectAll
120                   End If
130               Else
140                   If block.left + block.width >= left Then
150                       selectStart = True
160                       block.selectSome left, right
                      
170                       If right = -1 Then
                              Dim count2 As Integer
                          
180                           For count2 = count + 1 To m_blocks.count
190                               Set block = m_blocks.item(count2)
200                               block.selectAll
210                           Next count2
                              
220                           Exit For
230                       End If
                      
240                       If right <= block.left + block.width Then
250                           selectEnd = True
260                       End If
270                   Else
280                       If block.selected Then block.unSelect
290                   End If
300               End If
310           End If
320       Next count
          
330       m_selected = True
End Sub

Public Sub selectAll()
          Dim count As Integer
          Dim block As ITextRenderBlock
          
10        For count = 1 To m_blocks.count
20            Set block = m_blocks.item(count)
30            block.selectAll
40        Next count
          
50        m_selected = True
End Sub

Public Sub unSelect()
          Dim count As Integer
          Dim block As ITextRenderBlock

10        For count = 1 To m_blocks.count
20            Set block = m_blocks.item(count)
30            block.unSelect
40        Next count
          
50        m_selected = False
End Sub

Public Function getSelectedText(drawingData As CDrawingData, withFormatCodes As Boolean) As String
          Dim count As Integer
          Dim block As ITextRenderBlock
          
          Dim text As String
          
10        For count = 1 To m_blocks.count
20            Set block = m_blocks.item(count)
30            text = text & block.getSelectedText(drawingData, withFormatCodes)
40        Next count
          
50        getSelectedText = text
End Function

Friend Function getMouseOverBlock(ByVal x As Long) As ITextRenderBlock
          Dim count As Long
          Dim block As ITextRenderBlock
          
10        For count = 1 To m_blocks.count
20            Set block = m_blocks.item(count)
              
30            If block.width <> 0 Then
40                If x >= m_blocks.item(count).left Then
50                    If x <= m_blocks.item(count).left + m_blocks.item(count).width Then
60                        Set getMouseOverBlock = block
70                        Exit Function
80                    End If
90                End If
100           End If
110       Next count
End Function

Public Sub render(drawingData As CDrawingData)
          Dim count As Integer
          Dim block As ITextRenderBlock
          
10        For count = 1 To m_blocks.count
20            Set block = m_blocks.item(count)
              
30            If block.width <> 0 Then
40                Exit For
50            End If
60        Next count
          
70        If block Is Nothing Then
80            Exit Sub
90        End If
          
100       If block.left > 0 Then
             'Indented
110           drawingData.fillLineIndent block.left
120       End If
          
130       For count = 1 To m_blocks.count
140           Set block = m_blocks.item(count)
150           block.render drawingData
160       Next count
          
170       m_top = drawingData.realY - drawingData.fontHeight
180       m_bottom = drawingData.realY
          
190       drawingData.fillRemainingLine
End Sub
