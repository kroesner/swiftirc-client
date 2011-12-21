VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum eLineSep
    elsNone
    elsTop
    elsBottom
    elsBoth
End Enum

Private m_text As String
Private m_physLines As New cArrayList
Private m_oldPhysLines As Long
Private m_wrapped As Boolean
Private m_top As Integer
Private m_bottom As Integer
Private m_height As Integer
Private m_selected As Boolean

Private m_eventDef As CEvent
Private m_userStyle As CUserStyle
Private m_username As String

Private m_timeStamp As Long

'Is all or part of the line currently on screen
Private m_visible As Boolean

Private m_wasDisplayed As Boolean

Private m_format As String
Private m_flags As Long
Private m_foreColour As Byte
Private m_eventColour As CEventColour
Private m_params() As String

Private m_seperatorAbove As Boolean
Private m_seperatorBelow As Boolean

Private m_toggleBold As Boolean
Private m_toggleItalic As Boolean
Private m_toggleReverse As Boolean
Private m_toggleUnderline As Boolean

Public Sub init(aEvent As CEvent, flags As Long, userStyle As CUserStyle, username As String, _
    params() As String)
    Set m_eventDef = aEvent
    m_flags = flags
    Set m_userStyle = userStyle
    m_username = username
    m_params = params
End Sub

Public Sub initEx(eventColour As CEventColour, foreColour As Byte, format As String, userStyle As _
    CUserStyle, username As String, flags As Long, params() As String)

    Set m_eventColour = eventColour
    m_foreColour = foreColour
    m_format = format
    Set m_userStyle = userStyle
    m_username = username
    m_flags = flags
    m_params = params
End Sub

Public Property Get param(index As Long) As String
    param = m_params(index)
End Property

Public Property Get shouldShow() As Boolean
    If Not m_eventDef Is Nothing Then
        shouldShow = m_eventDef.visible
    Else
        shouldShow = True
    End If
End Property

Public Property Get wasDisplayed() As Boolean
    wasDisplayed = m_wasDisplayed
End Property

Public Property Let wasDisplayed(newValue As Boolean)
    m_wasDisplayed = newValue
End Property

Public Property Get selected() As Boolean
    selected = m_selected
End Property

Public Property Get visible() As Boolean
    visible = m_visible
End Property

Public Property Let visible(newValue As Boolean)
    m_visible = newValue
End Property

Private Property Get timestamps() As Boolean
    If g_timestamps = False Then
        Exit Property
    End If

    If Not m_flags And TVE_NOEVENT Then
        timestamps = m_eventDef.flags And TVE_TIMESTAMP
    End If
    
    If Not timestamps Then
        timestamps = m_flags And TVE_TIMESTAMP
    End If
End Property

Private Property Get indent() As Boolean
    If Not m_flags And TVE_NOEVENT Then
        indent = m_eventDef.flags And TVE_INDENTWRAP
    End If
    
    If Not indent Then
        indent = m_flags And TVE_INDENTWRAP
    End If
End Property

Public Property Get seperatorAbove() As Boolean
    seperatorAbove = m_seperatorAbove
End Property

Public Property Let seperatorAbove(newValue As Boolean)
    m_seperatorAbove = newValue
End Property

Public Property Get seperatorBelow() As Boolean
    seperatorBelow = m_seperatorBelow
End Property

Public Property Let seperatorBelow(newValue As Boolean)
    m_seperatorBelow = newValue
End Property

Public Function seperatorTop(ignoreSeperators As Boolean) As Boolean
    If Not m_flags And TVE_NOEVENT Then
        If m_eventDef.flags And TVE_SEPERATE_TOP Then
            If Not ignoreSeperators Or m_eventDef.flags And TVE_SEPERATE_EXPLICIT Then
                seperatorTop = True
            End If
            
            Exit Function
        End If
    End If
    
    If m_flags And TVE_SEPERATE_TOP Then
        If Not ignoreSeperators Or m_flags And TVE_SEPERATE_EXPLICIT Then
            seperatorTop = True
        End If
    End If
End Function

Public Function seperatorBottom(ignoreSeperators As Boolean) As Boolean
    If Not m_flags And TVE_NOEVENT Then
        If m_eventDef.flags And TVE_SEPERATE_BOTTOM Then
            If Not ignoreSeperators Or m_eventDef.flags And TVE_SEPERATE_EXPLICIT Then
                seperatorBottom = True
            End If
            
            Exit Function
        End If
    End If
    
    If m_flags And TVE_SEPERATE_BOTTOM Then
        If Not ignoreSeperators Or m_flags And TVE_SEPERATE_EXPLICIT Then
            seperatorBottom = True
        End If
    End If
End Function

Private Property Get lineSeperateType() As eLineSep
    If Not m_flags And TVE_NOEVENT Then
        If m_eventDef.flags And TVE_SEPERATE_BOTH Then
            lineSeperateType = elsBoth
        ElseIf m_eventDef.flags And TVE_SEPERATE_TOP Then
            lineSeperateType = elsTop
        ElseIf m_eventDef.flags And TVE_SEPERATE_BOTTOM Then
            lineSeperateType = elsBottom
        Else
            lineSeperateType = elsNone
        End If
    End If
    
    If lineSeperateType = elsNone Then
        If m_flags And TVE_SEPERATE_BOTH Then
            lineSeperateType = elsBoth
        ElseIf m_flags And TVE_SEPERATE_TOP Then
            lineSeperateType = elsTop
        ElseIf m_flags And TVE_SEPERATE_BOTTOM Then
            lineSeperateType = elsBottom
        Else
            lineSeperateType = elsNone
        End If
    End If
End Property

Public Sub setSelection(top As Integer, bottom As Integer, left As Integer, right As Integer, _
    clickX As Integer, clickY As Integer, moveX As Integer)
    
    Dim count As Integer
    Dim line As CPhysLine
    
    Dim clickedLine As Integer
    
    If clickY > top Then
        For count = 1 To m_physLines.count
            If m_physLines.item(count).bottom >= clickY Then
                clickedLine = count
                Exit For
            End If
        Next count
    Else
        For count = m_physLines.count To 1 Step -1
            If m_physLines.item(count).top <= clickY Then
                clickedLine = count
                Exit For
            End If
        Next count
    End If
    
    For count = 1 To m_physLines.count
        Set line = m_physLines.item(count)
    
        If line.bottom >= top And line.top <= bottom Then
            If line.top < top Then
                If line.bottom > bottom Then
                    line.setSelection left, right
                Else
                    If count = clickedLine Then
                        line.setSelection clickX, -1
                    Else
                        line.setSelection moveX, -1
                    End If
                End If
            ElseIf line.bottom > bottom Then
                If line.top > top Then
                    If count = clickedLine Then
                        line.setSelection 0, clickX
                    Else
                        line.setSelection 0, moveX
                    End If
                Else
                    line.setSelection left, right
                End If
            Else
                line.selectAll
            End If
        Else
            If m_physLines.item(count).selected Then
                m_physLines.item(count).unSelect
            End If
        End If
    Next count
    
    m_selected = True
End Sub

Public Sub unSelect()
    Dim count As Integer
    
    For count = 1 To m_physLines.count
        If m_physLines.item(count).selected Then
            m_physLines.item(count).unSelect
        End If
    Next count
    
    m_selected = False
End Sub

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

Public Property Get height() As Integer
    height = m_height
End Property

Public Property Let height(newValue As Integer)
    m_height = newValue
End Property

Public Property Get text() As String
    text = m_text
End Property

Public Property Let text(newText As String)
    m_text = newText
End Property

Public Property Get physLineCount() As Long
    If m_physLines.count > 0 Then
        physLineCount = m_physLines.count
    Else
        physLineCount = m_oldPhysLines
    End If
End Property

Public Property Get realPhysLineCount() As Long
    realPhysLineCount = m_physLines.count
End Property

Public Property Get physLine(index As Long)
    If index > 0 And index <= m_physLines.count Then
        Set physLine = m_physLines.item(index)
    End If
End Property

Public Property Get wrapped() As Boolean
    wrapped = m_wrapped
End Property

Private Function addPhysLine() As CPhysLine
    Set addPhysLine = New CPhysLine
    m_physLines.Add addPhysLine
End Function

Private Function isFormatBlock(block As ITextRenderBlock) As Boolean
    isFormatBlock = True

    If TypeOf block Is CBlockBold Then
    ElseIf TypeOf block Is CBlockItalic Then
    ElseIf TypeOf block Is CBlockUnderline Then
    ElseIf TypeOf block Is CBlockForeColour Then
    ElseIf TypeOf block Is CBlockBackColour Then
    Else
        isFormatBlock = False
    End If
End Function

Private Function isFormatCode(wChar As Integer) As Boolean
    isFormatCode = True

    Select Case wChar
        Case 2
        Case 3
        Case 4
        Case 15
        Case 22
        Case 31
        Case Else
            isFormatCode = False
    End Select
End Function

Public Function getText(codes As Boolean, includeTimestamp As Boolean, ignoreSeperators As Boolean) As String
    Dim count As Long
    Dim text As String
    Dim format As String
    Dim wChar As Long
    
   On Error GoTo getText_Error

    If m_flags And TVE_NOEVENT Then
        format = m_format
    Else
        format = m_eventDef.format
    End If
    
    For count = 1 To Len(format)
        wChar = AscW(Mid$(format, count, 1))
        If wChar = 36 Then
            wChar = AscW(Mid$(format, count + 1, 1))
            
            If wChar > 47 And wChar < 58 Then
                If UBound(m_params) >= wChar - 48 Then
                    text = text & m_params(wChar - 48)
                End If
                
                count = count + 1
            Else
                Select Case wChar
                    Case AscW("s")
                        If Not m_userStyle Is Nothing Then
                            If Not m_userStyle.image Is Nothing Then
                                text = text & m_userStyle.image.caption
                            End If
                        End If
                        
                        text = text & m_username
                        count = count + 1
                    Case 111
                End Select
            End If
        Else
            text = text & ChrW$(wChar)
        End If
    Next count

    Dim text2 As String
    
    If seperatorTop(ignoreSeperators) And Not seperatorAbove Then
        text2 = "-" & vbCrLf
    End If
    
    If includeTimestamp Then
        If timestamps Then
            Dim timestamp As CBlockText
        
            Set timestamp = renderTimestamp
            text2 = text2 & timestamp.text
        End If
    End If
    
    If Not codes Then
        For count = 1 To Len(text)
            wChar = AscW(Mid$(text, count, 1))
            
            Select Case wChar
                Case 2
                Case 4
                Case 15
                Case 22
                Case 31
                Case 3
                    If count = Len(text) Then
                        Exit For
                    End If
                    
                    Dim foreColour As Byte
                    Dim backColour As Byte
                    
                    count = parseColourCode(text, count + 1, foreColour, backColour) - 1
                Case Else
                    text2 = text2 & ChrW$(wChar)
            End Select
        Next count
    Else
        text2 = text2 & text
    End If
    
    If seperatorBottom(ignoreSeperators) And Not seperatorBelow Then
        text2 = text2 & vbCrLf & "-" & vbCrLf
    Else
        text2 = text2 & vbCrLf
    End If
    
    getText = text2

   On Error GoTo 0
   Exit Function

getText_Error:
    handleError "getText", Err.Number, Err.Description, Erl, "Format: " & format & " Text: " & text
End Function

Private Sub parseIrcFormatting(text As String, allowPlaceHolders As Boolean, ByRef blocks As _
    cArrayList, Optional allowArtifacts As Boolean = True)
    
    Dim count As Long
    Dim wChar As Integer
    Dim last As Long
    Dim Length As Long
    
    Dim blockText As CBlockText
    Dim blockBold As CBlockBold
    Dim blockItalic As CBlockItalic
    Dim blockUnderline As CBlockUnderline
    Dim blockReverse As CBlockReverse
    Dim blockForeColour As CBlockForeColour
    Dim blockbackColour As CBlockBackColour
    Dim blockImage As CBlockImage
    
    Dim isUrl As Boolean
    
    last = 1
    
    For count = 1 To Len(text)
        wChar = AscW(Mid$(text, count, 1))
        
        If isFormatCode(wChar) Or wChar = 36 And allowPlaceHolders Then
            Length = count - last
            
            If Length > 0 Then
                Set blockText = New CBlockText
                
                blockText.text = Mid$(text, last, Length)
                blocks.Add blockText
            End If
            
            Select Case wChar
                Case 36
                    If allowPlaceHolders Then
                        wChar = Asc(Mid$(text, count + 1, 1))
                        
                        If wChar > 47 And wChar < 58 Then
                            If UBound(m_params) >= wChar - 48 Then
                                parseIrcFormatting m_params(wChar - 48), False, blocks
                            End If
                            
                            count = count + 1
                        Else
                            Select Case wChar
                                Case AscW("s")
                                    If Not m_userStyle Is Nothing Then
                                        m_userStyle.buildFormat blocks
                                    End If
                                    
                                    Set blockText = New CBlockText
                                    blockText.text = m_username
                                    blocks.Add blockText
                                    blocks.Add New CBlockNormal
                                    
                                    count = count + 1
                                Case 111
                                    blocks.Add New CBlockNormal
                                    count = count + 1
                            End Select
                        End If
                    End If
                Case 2
                    m_toggleBold = Not m_toggleBold
                    Set blockBold = New CBlockBold
                    blockBold.bold = m_toggleBold
                    blocks.Add blockBold
                Case 4
                    m_toggleItalic = Not m_toggleItalic
                    Set blockItalic = New CBlockItalic
                    blockItalic.italic = m_toggleItalic
                    blocks.Add blockItalic
                Case 15
                    m_toggleBold = False
                    m_toggleItalic = False
                    m_toggleUnderline = False
                    m_toggleReverse = False
                    
                    blocks.Add New CBlockNormal
                Case 22
                    m_toggleReverse = Not m_toggleReverse
                    Set blockReverse = New CBlockReverse
                    blockReverse.reverse = m_toggleReverse
                    blocks.Add blockReverse
                Case 31
                    m_toggleUnderline = Not m_toggleUnderline
                    Set blockUnderline = New CBlockUnderline
                    blockUnderline.underline = m_toggleUnderline
                    blocks.Add blockUnderline
                Case 3
                    If count = Len(text) Then
                        blocks.Add New CBlockForeColour
                        blocks.Add New CBlockBackColour
                        Exit Sub
                    End If
                    
                    Dim foreColour As Byte
                    Dim backColour As Byte
                    
                    count = parseColourCode(text, count + 1, foreColour, backColour) - 1
                    
                    If foreColour <> 255 Then
                        Set blockForeColour = New CBlockForeColour
                        
                        blockForeColour.hasForeColour = True
                        blockForeColour.foreColour = foreColour
                        blocks.Add blockForeColour
                        
                    Else
                        blocks.Add New CBlockForeColour
                        
                        If backColour = 255 Then
                            blocks.Add New CBlockBackColour
                        End If
                    End If
                    
                    If backColour <> 255 Then
                        Set blockbackColour = New CBlockBackColour
                        
                        blockbackColour.hasBackColour = True
                        blockbackColour.backColour = backColour
                        blocks.Add blockbackColour
                    End If
            End Select
            
            last = count + 1
        Else
            isUrl = False
        
            If allowArtifacts Then
                If wChar = AscW("h") Or wChar = AscW("H") Then
                    If count < Len(text) Then
                        If Mid$(LCase$(text), count, 7) = "http://" Then
                            isUrl = True
                        End If
                    End If
                    
                    If count < Len(text) Then
                        If Mid$(LCase$(text), count, 8) = "https://" Then
                            isUrl = True
                        End If
                    End If
                ElseIf wChar = AscW("w") Or wChar = AscW("W") Then
                    If count < Len(text) Then
                        If Mid$(LCase$(text), count, 4) = "www." Then
                            isUrl = True
                        End If
                    End If
                ElseIf wChar = AscW("#") Then
                    isUrl = True
                End If
            End If
            
            If isUrl Then
                Length = count - last
                
                If Length > 0 Then
                    Set blockText = New CBlockText
    
                    blockText.text = Mid$(text, last, Length)
                    blocks.Add blockText
                End If
                
                count = count + parseLink(text, count, blocks)
                last = count
            End If
        End If
    Next count
    
    If count - last > 0 Then
        If last <= Len(text) Then
            Set blockText = New CBlockText
        
            blockText.text = Mid$(text, last)
            blocks.Add blockText
        End If
    End If
End Sub

Private Function parseLink(ByRef text As String, index As Long, blocks As cArrayList) As Long
    Dim url As String
    Dim count As Long
    
    For count = index To Len(text)
        If Mid$(text, count, 1) = "'" Then
            url = Mid$(text, index, count - index)
            Exit For
        ElseIf Mid$(text, count, 1) = "]" Then
            url = Mid$(text, index, count - index)
            Exit For
        ElseIf Mid$(text, count, 1) = """" Then
            url = Mid$(text, index, count - index)
            Exit For
        ElseIf Mid$(text, count, 1) = ")" Then
            url = Mid$(text, index, count - index)
            Exit For
        ElseIf Mid$(text, count, 1) = "." Then
            If count < Len(text) Then
                If Mid$(text, count + 1, 1) = " " Then
                    url = Mid$(text, index, count - index)
                    Exit For
                End If
            Else
                url = Mid$(text, index, count - index)
                Exit For
            End If
        ElseIf Mid$(text, count, 1) = " " Then
            url = Mid$(text, index, count - index)
            Exit For
        End If
    Next count
    
    If LenB(url) = 0 Then
        url = Mid$(text, index)
    End If
    
    Dim urlBlock As New CBlockText

    urlBlock.isUrl = True
    blocks.Add urlBlock
    
    Dim urlBlocks As New cArrayList
    Dim block As ITextRenderBlock
    Dim blockText As CBlockText
    
    parseIrcFormatting url, False, urlBlocks, False
    
    For count = 1 To urlBlocks.count
        Set block = urlBlocks.item(count)
    
        If TypeOf block Is CBlockText Then
            Set blockText = block
            urlBlock.text = urlBlock.text & blockText.text
        Else
            blocks.Add block
        End If
    Next count
    
    urlBlock.url = urlBlock.text
    
    parseLink = Len(url)
End Function

Private Function parseColourCode(text As String, start As Integer, ByRef fore As Byte, ByRef back _
    As Byte) As Integer
    
    Dim colourCount As Integer
    Dim digits As Byte
    Dim currentColour As Byte
    Dim hasColour As Boolean
    
    Dim wChar As Integer
    
   On Error GoTo parseColourCode_Error

    fore = 255
    back = 255
    
    For colourCount = start To Len(text)
        wChar = AscW(Mid$(text, colourCount, 1))
        
        If wChar > 47 And wChar < 58 Then
            If digits = 0 Then
                hasColour = True
                currentColour = wChar - 48
                digits = 1
                start = start + 1
            ElseIf digits = 1 Then
                currentColour = (currentColour * 10) + (wChar - 48)
                digits = 2
                start = start + 1
            Else
                Exit For
            End If
        ElseIf wChar = AscW(",") Then
            If Not hasColour Then
                Exit For
            End If
            
            fore = currentColour
            
            If fore > 15 Then
                fore = fore Mod 16
            End If
            
            hasColour = False
            digits = 0
            start = start + 1
        Else
            Exit For
        End If
    Next colourCount

    If fore <> 255 Then
        If hasColour Then
            back = currentColour
            
            If back > 15 Then
                back = back Mod 16
            End If
        End If
    ElseIf hasColour Then
        fore = currentColour
        
        If fore > 15 Then
            fore = fore Mod 16
        End If
    End If
    
    parseColourCode = start

   On Error GoTo 0
   Exit Function

parseColourCode_Error:
    handleError "parseColourCode", Err.Number, Err.Description, Erl, vbNullString
End Function

Private Function renderTimestamp() As CBlockText
    Set renderTimestamp = New CBlockText
    
    Dim aDate As Date
    
    aDate = CDate(25569 + (m_timeStamp / 86400))
    
    renderTimestamp.text = swiftFormatTime(aDate, g_timestampFormat) & " "
End Function

Private Sub addSeperator(drawingData As CDrawingData)
    Dim physLine As CPhysLine
    Dim seperator As New CBlockText
    Dim textSize As SIZE
    
    swiftGetTextExtentPoint32 drawingData.Dc, "-", textSize
    
    seperator.text = "-"
    seperator.left = 0
    seperator.width = textSize.cx
    
    Set physLine = addPhysLine
    physLine.addBlock seperator
End Sub

Public Function wordWrap(drawingData As CDrawingData) As Long
    If m_wrapped Then
        Exit Function
    End If
    
    m_physLines.clear
    
    Dim maxFit As Long
    Dim textSize As SIZE

    Dim break As Integer
    Dim blocks As New cArrayList

    If timestamps Then blocks.Add renderTimestamp

    If m_flags And TVE_NOEVENT Then
        parseIrcFormatting m_format, True, blocks
    Else
        parseIrcFormatting m_eventDef.format, True, blocks
    End If
    
    Dim blockCount As Integer
    
    Dim block As ITextRenderBlock
    
    Dim textBlock As CBlockText
    Dim newTextBlock As CBlockText
    
    Dim oldFont As Long
    
    Dim width As Long
    Dim x As Integer
    
    width = drawingData.width
    
    drawingData.bold = False
    drawingData.italic = False
    drawingData.underline = False
    m_toggleBold = False
    m_toggleItalic = False
    m_toggleUnderline = False
    m_toggleReverse = False
    
    If seperatorTop(drawingData.ignoreSeperators) And Not seperatorAbove Then
        addSeperator drawingData
    End If
    
    Dim tooSmall As Boolean
    Dim noBreakTooSmall As Boolean
    Dim physLine As CPhysLine

    Set physLine = addPhysLine
    
    Do While blockCount < blocks.count
        blockCount = blockCount + 1
        Set block = blocks.item(blockCount)
        
        If TypeOf block Is CBlockText Then
            Set textBlock = block
        
            swiftGetTextExtentExPoint drawingData.Dc, textBlock.text, width, maxFit, textSize
        
            If Len(textBlock.text) > maxFit Then
                If maxFit <> 0 Then
                    tooSmall = False
                
                    For break = maxFit To 1 Step -1
                        If AscW(Mid$(textBlock.text, break, 1)) = 32 Then
                            Exit For
                        End If
                    Next break
                    
                    If break = 0 Then
                        If Not noBreakTooSmall And physLine.blockCount > 0 Then
                            blockCount = blockCount - 1
                            noBreakTooSmall = True
                        Else
                            break = maxFit
                        End If
                    Else
                        noBreakTooSmall = False
                    End If
                    
                    If break > 0 Then
                        Set newTextBlock = New CBlockText
                        
                        newTextBlock.text = Mid$(textBlock.text, break + 1)
                        newTextBlock.isUrl = textBlock.isUrl
                        newTextBlock.url = textBlock.url
                        
                        blocks.Add newTextBlock, blockCount + 1
                        
                        textBlock.text = Mid$(textBlock.text, 1, break)
                        
                        textBlock.left = x
                        
                        swiftGetTextExtentPoint32 drawingData.Dc, textBlock.text, _
                            textSize
                        textBlock.width = textSize.cx
                        
                        physLine.addBlock textBlock
                    End If
                Else
                    If tooSmall Then
                        Exit Do
                    End If
                    
                    tooSmall = True
                    blockCount = blockCount - 1
                End If
                
                Set physLine = addPhysLine
                
                If indent Then
                    width = drawingData.width - 10
                    x = 10
                Else
                    width = drawingData.width
                    x = 0
                End If
                
                physLine.cloneFormat drawingData
            Else
                tooSmall = False
                noBreakTooSmall = False
                textBlock.left = x
                textBlock.width = textSize.cx
                
                physLine.addBlock textBlock
                
                x = x + textSize.cx
                width = width - textSize.cx
            End If
        ElseIf TypeOf block Is CBlockImage Then
            tooSmall = False
            
            If width < drawingData.fontHeight Then
                Set physLine = addPhysLine
                x = drawingData.fontHeight
                width = drawingData.width - drawingData.fontHeight
                
                block.left = 0
                block.width = drawingData.fontHeight
                
                physLine.addBlock block
            Else
                block.left = x
                block.width = drawingData.fontHeight
                width = width - drawingData.fontHeight
                x = x + drawingData.fontHeight
                physLine.addBlock block
            End If
        Else
            tooSmall = False
        
            block.render drawingData
            physLine.addBlock block
        End If
    Loop
    
    If seperatorBottom(drawingData.ignoreSeperators) And Not seperatorBelow Then
        addSeperator drawingData
    End If
    
    wordWrap = m_physLines.count - m_oldPhysLines
    
    m_oldPhysLines = m_physLines.count
    m_wrapped = True
End Function

Public Sub clearWrap()
    m_physLines.clear
    m_wrapped = False
End Sub

Public Sub needsWrapping()
    m_wrapped = False
End Sub

Public Sub render(drawingData As CDrawingData, physLine As Long)
    Dim count As Integer
    Dim start As Integer
    
    If physLine > 0 Then
        start = physLine
    Else
        start = m_physLines.count
    End If
    
    Dim line As CPhysLine
    
    Dim blockRect As RECT
    
    
    If m_flags And TVE_NOEVENT Then
        If m_flags And TVE_CUSTOMIRCCOLOUR Then
            drawingData.lineForeColour = m_foreColour
        Else
            drawingData.lineForeColour = m_eventColour.colour
        End If
    Else
        drawingData.lineForeColour = m_eventDef.eventColour.colour
    End If

    m_bottom = drawingData.realY
    
    drawingData.reset
    
    For count = start To 1 Step -1
        drawingData.x = 0
    
        Set line = m_physLines.item(count)
        line.render drawingData
        drawingData.reset
        
        drawingData.y = drawingData.y - drawingData.fontHeight
        drawingData.realY = drawingData.realY - drawingData.fontHeight
        
        If drawingData.y <= -drawingData.fontHeight Then
            Exit For
        End If
    Next count

    m_top = drawingData.realY
End Sub

Public Function getSelectedText(drawingData As CDrawingData, withFormatCodes As Boolean) As String
    Dim count As Integer
    Dim text As String
    
    For count = 1 To m_physLines.count
        text = text & m_physLines.item(count).getSelectedText(drawingData, withFormatCodes)
    Next count
    
    text = text & vbCrLf
    
    getSelectedText = text
End Function

Public Sub shiftedUp(pixels As Long)
    m_top = m_top - pixels
    m_bottom = m_bottom - pixels
    
    Dim count As Integer
    
    For count = 1 To m_physLines.count
        m_physLines.item(count).top = m_physLines.item(count).top - pixels
        m_physLines.item(count).bottom = m_physLines.item(count).bottom - pixels
    Next count
End Sub

Public Sub shiftedDown(pixels As Long)
    m_top = m_top + pixels
    m_bottom = m_bottom + pixels
    
    Dim count As Integer
    
    For count = 1 To m_physLines.count
        m_physLines.item(count).top = m_physLines.item(count).top + pixels
        m_physLines.item(count).bottom = m_physLines.item(count).bottom + pixels
    Next count
End Sub

Private Sub Class_Initialize()
    m_oldPhysLines = 1
    m_timeStamp = CLng(DateDiff("s", #1/1/1970#, Now))
End Sub
