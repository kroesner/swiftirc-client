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
10        Set m_eventDef = aEvent
20        m_flags = flags
30        Set m_userStyle = userStyle
40        m_username = username
50        m_params = params
End Sub

Public Sub initEx(eventColour As CEventColour, foreColour As Byte, format As String, userStyle As _
    CUserStyle, username As String, flags As Long, params() As String)

10        Set m_eventColour = eventColour
20        m_foreColour = foreColour
30        m_format = format
40        Set m_userStyle = userStyle
50        m_username = username
60        m_flags = flags
70        m_params = params
End Sub

Public Property Get param(index As Long) As String
10        param = m_params(index)
End Property

Public Property Get shouldShow() As Boolean
10        If Not m_eventDef Is Nothing Then
20            shouldShow = m_eventDef.visible
30        Else
40            shouldShow = True
50        End If
End Property

Public Property Get wasDisplayed() As Boolean
10        wasDisplayed = m_wasDisplayed
End Property

Public Property Let wasDisplayed(newValue As Boolean)
10        m_wasDisplayed = newValue
End Property

Public Property Get selected() As Boolean
10        selected = m_selected
End Property

Public Property Get visible() As Boolean
10        visible = m_visible
End Property

Public Property Let visible(newValue As Boolean)
10        m_visible = newValue
End Property

Private Property Get timestamps() As Boolean
10        If g_timestamps = False Then
20            Exit Property
30        End If

40        If Not m_flags And TVE_NOEVENT Then
50            timestamps = m_eventDef.flags And TVE_TIMESTAMP
60        End If
          
70        If Not timestamps Then
80            timestamps = m_flags And TVE_TIMESTAMP
90        End If
End Property

Private Property Get indent() As Boolean
10        If Not m_flags And TVE_NOEVENT Then
20            indent = m_eventDef.flags And TVE_INDENTWRAP
30        End If
          
40        If Not indent Then
50            indent = m_flags And TVE_INDENTWRAP
60        End If
End Property

Public Property Get seperatorAbove() As Boolean
10        seperatorAbove = m_seperatorAbove
End Property

Public Property Let seperatorAbove(newValue As Boolean)
10        m_seperatorAbove = newValue
End Property

Public Property Get seperatorBelow() As Boolean
10        seperatorBelow = m_seperatorBelow
End Property

Public Property Let seperatorBelow(newValue As Boolean)
10        m_seperatorBelow = newValue
End Property

Public Function seperatorTop(ignoreSeperators As Boolean) As Boolean
10        If Not m_flags And TVE_NOEVENT Then
20            If m_eventDef.flags And TVE_SEPERATE_TOP Then
30                If Not ignoreSeperators Or m_eventDef.flags And TVE_SEPERATE_EXPLICIT Then
40                    seperatorTop = True
50                End If
                  
60                Exit Function
70            End If
80        End If
          
90        If m_flags And TVE_SEPERATE_TOP Then
100           If Not ignoreSeperators Or m_flags And TVE_SEPERATE_EXPLICIT Then
110               seperatorTop = True
120           End If
130       End If
End Function

Public Function seperatorBottom(ignoreSeperators As Boolean) As Boolean
10        If Not m_flags And TVE_NOEVENT Then
20            If m_eventDef.flags And TVE_SEPERATE_BOTTOM Then
30                If Not ignoreSeperators Or m_eventDef.flags And TVE_SEPERATE_EXPLICIT Then
40                    seperatorBottom = True
50                End If
                  
60                Exit Function
70            End If
80        End If
          
90        If m_flags And TVE_SEPERATE_BOTTOM Then
100           If Not ignoreSeperators Or m_flags And TVE_SEPERATE_EXPLICIT Then
110               seperatorBottom = True
120           End If
130       End If
End Function

Private Property Get lineSeperateType() As eLineSep
10        If Not m_flags And TVE_NOEVENT Then
20            If m_eventDef.flags And TVE_SEPERATE_BOTH Then
30                lineSeperateType = elsBoth
40            ElseIf m_eventDef.flags And TVE_SEPERATE_TOP Then
50                lineSeperateType = elsTop
60            ElseIf m_eventDef.flags And TVE_SEPERATE_BOTTOM Then
70                lineSeperateType = elsBottom
80            Else
90                lineSeperateType = elsNone
100           End If
110       End If
          
120       If lineSeperateType = elsNone Then
130           If m_flags And TVE_SEPERATE_BOTH Then
140               lineSeperateType = elsBoth
150           ElseIf m_flags And TVE_SEPERATE_TOP Then
160               lineSeperateType = elsTop
170           ElseIf m_flags And TVE_SEPERATE_BOTTOM Then
180               lineSeperateType = elsBottom
190           Else
200               lineSeperateType = elsNone
210           End If
220       End If
End Property

Public Sub setSelection(top As Integer, bottom As Integer, left As Integer, right As Integer, _
    clickX As Integer, clickY As Integer, moveX As Integer)
          
          Dim count As Integer
          Dim line As CPhysLine
          
          Dim clickedLine As Integer
          
10        If clickY > top Then
20            For count = 1 To m_physLines.count
30                If m_physLines.item(count).bottom >= clickY Then
40                    clickedLine = count
50                    Exit For
60                End If
70            Next count
80        Else
90            For count = m_physLines.count To 1 Step -1
100               If m_physLines.item(count).top <= clickY Then
110                   clickedLine = count
120                   Exit For
130               End If
140           Next count
150       End If
          
160       For count = 1 To m_physLines.count
170           Set line = m_physLines.item(count)
          
180           If line.bottom >= top And line.top <= bottom Then
190               If line.top < top Then
200                   If line.bottom > bottom Then
210                       line.setSelection left, right
220                   Else
230                       If count = clickedLine Then
240                           line.setSelection clickX, -1
250                       Else
260                           line.setSelection moveX, -1
270                       End If
280                   End If
290               ElseIf line.bottom > bottom Then
300                   If line.top > top Then
310                       If count = clickedLine Then
320                           line.setSelection 0, clickX
330                       Else
340                           line.setSelection 0, moveX
350                       End If
360                   Else
370                       line.setSelection left, right
380                   End If
390               Else
400                   line.selectAll
410               End If
420           Else
430               If m_physLines.item(count).selected Then
440                   m_physLines.item(count).unSelect
450               End If
460           End If
470       Next count
          
480       m_selected = True
End Sub

Public Sub unSelect()
          Dim count As Integer
          
10        For count = 1 To m_physLines.count
20            If m_physLines.item(count).selected Then
30                m_physLines.item(count).unSelect
40            End If
50        Next count
          
60        m_selected = False
End Sub

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

Public Property Get height() As Integer
10        height = m_height
End Property

Public Property Let height(newValue As Integer)
10        m_height = newValue
End Property

Public Property Get text() As String
10        text = m_text
End Property

Public Property Let text(newText As String)
10        m_text = newText
End Property

Public Property Get physLineCount() As Long
10        If m_physLines.count > 0 Then
20            physLineCount = m_physLines.count
30        Else
40            physLineCount = m_oldPhysLines
50        End If
End Property

Public Property Get realPhysLineCount() As Long
10        realPhysLineCount = m_physLines.count
End Property

Public Property Get physLine(index As Long)
10        If index > 0 And index <= m_physLines.count Then
20            Set physLine = m_physLines.item(index)
30        End If
End Property

Public Property Get wrapped() As Boolean
10        wrapped = m_wrapped
End Property

Private Function addPhysLine() As CPhysLine
10        Set addPhysLine = New CPhysLine
20        m_physLines.Add addPhysLine
End Function

Private Function isFormatBlock(block As ITextRenderBlock) As Boolean
10        isFormatBlock = True

20        If TypeOf block Is CBlockBold Then
30        ElseIf TypeOf block Is CBlockItalic Then
40        ElseIf TypeOf block Is CBlockUnderline Then
50        ElseIf TypeOf block Is CBlockForeColour Then
60        ElseIf TypeOf block Is CBlockBackColour Then
70        Else
80            isFormatBlock = False
90        End If
End Function

Private Function isFormatCode(wChar As Integer) As Boolean
10        isFormatCode = True

20        Select Case wChar
              Case 2
30            Case 3
40            Case 4
50            Case 15
60            Case 22
70            Case 31
80            Case Else
90                isFormatCode = False
100       End Select
End Function

Public Function getText(codes As Boolean, includeTimestamp As Boolean, ignoreSeperators As Boolean) As String
          Dim count As Long
          Dim text As String
          Dim format As String
          Dim wChar As Long
          
10       On Error GoTo getText_Error

20        If m_flags And TVE_NOEVENT Then
30            format = m_format
40        Else
50            format = m_eventDef.format
60        End If
          
70        For count = 1 To Len(format)
80            wChar = AscW(Mid$(format, count, 1))
90            If wChar = 36 Then
100               wChar = AscW(Mid$(format, count + 1, 1))
                  
110               If wChar > 47 And wChar < 58 Then
120                   If UBound(m_params) >= wChar - 48 Then
130                       text = text & m_params(wChar - 48)
140                   End If
                      
150                   count = count + 1
160               Else
170                   Select Case wChar
                          Case AscW("s")
180                           If Not m_userStyle Is Nothing Then
190                               If Not m_userStyle.image Is Nothing Then
200                                   text = text & m_userStyle.image.caption
210                               End If
220                           End If
                              
230                           text = text & m_username
240                           count = count + 1
250                       Case 111
260                   End Select
270               End If
280           Else
290               text = text & ChrW$(wChar)
300           End If
310       Next count

          Dim text2 As String
          
320       If seperatorTop(ignoreSeperators) And Not seperatorAbove Then
330           text2 = "-" & vbCrLf
340       End If
          
350       If includeTimestamp Then
360           If timestamps Then
                  Dim timestamp As CBlockText
              
370               Set timestamp = renderTimestamp
380               text2 = text2 & timestamp.text
390           End If
400       End If
          
410       If Not codes Then
420           For count = 1 To Len(text)
430               wChar = AscW(Mid$(text, count, 1))
                  
440               Select Case wChar
                      Case 2
450                   Case 4
460                   Case 15
470                   Case 22
480                   Case 31
490                   Case 3
500                       If count = Len(text) Then
510                           Exit For
520                       End If
                          
                          Dim foreColour As Byte
                          Dim backColour As Byte
                          
530                       count = parseColourCode(text, count + 1, foreColour, backColour) - 1
540                   Case Else
550                       text2 = text2 & ChrW$(wChar)
560               End Select
570           Next count
580       Else
590           text2 = text2 & text
600       End If
          
610       If seperatorBottom(ignoreSeperators) And Not seperatorBelow Then
620           text2 = text2 & vbCrLf & "-" & vbCrLf
630       Else
640           text2 = text2 & vbCrLf
650       End If
          
660       getText = text2

670      On Error GoTo 0
680      Exit Function

getText_Error:
690       handleError "getText", Err.Number, Err.Description, Erl, "Format: " & format & " Text: " & text
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
          
10        last = 1
          
20        For count = 1 To Len(text)
30            wChar = AscW(Mid$(text, count, 1))
              
40            If isFormatCode(wChar) Or wChar = 36 And allowPlaceHolders Then
50                Length = count - last
                  
60                If Length > 0 Then
70                    Set blockText = New CBlockText
                      
80                    blockText.text = Mid$(text, last, Length)
90                    blocks.Add blockText
100               End If
                  
110               Select Case wChar
                      Case 36
120                       If allowPlaceHolders Then
130                           wChar = Asc(Mid$(text, count + 1, 1))
                              
140                           If wChar > 47 And wChar < 58 Then
150                               If UBound(m_params) >= wChar - 48 Then
160                                   parseIrcFormatting m_params(wChar - 48), False, blocks
170                               End If
                                  
180                               count = count + 1
190                           Else
200                               Select Case wChar
                                      Case AscW("s")
210                                       If Not m_userStyle Is Nothing Then
220                                           m_userStyle.buildFormat blocks
230                                       End If
                                          
240                                       Set blockText = New CBlockText
250                                       blockText.text = m_username
260                                       blocks.Add blockText
270                                       blocks.Add New CBlockNormal
                                          
280                                       count = count + 1
290                                   Case 111
300                                       blocks.Add New CBlockNormal
310                                       count = count + 1
320                               End Select
330                           End If
340                       End If
350                   Case 2
360                       m_toggleBold = Not m_toggleBold
370                       Set blockBold = New CBlockBold
380                       blockBold.bold = m_toggleBold
390                       blocks.Add blockBold
400                   Case 4
410                       m_toggleItalic = Not m_toggleItalic
420                       Set blockItalic = New CBlockItalic
430                       blockItalic.italic = m_toggleItalic
440                       blocks.Add blockItalic
450                   Case 15
460                       m_toggleBold = False
470                       m_toggleItalic = False
480                       m_toggleUnderline = False
490                       m_toggleReverse = False
                          
500                       blocks.Add New CBlockNormal
510                   Case 22
520                       m_toggleReverse = Not m_toggleReverse
530                       Set blockReverse = New CBlockReverse
540                       blockReverse.reverse = m_toggleReverse
550                       blocks.Add blockReverse
560                   Case 31
570                       m_toggleUnderline = Not m_toggleUnderline
580                       Set blockUnderline = New CBlockUnderline
590                       blockUnderline.underline = m_toggleUnderline
600                       blocks.Add blockUnderline
610                   Case 3
620                       If count = Len(text) Then
630                           blocks.Add New CBlockForeColour
640                           blocks.Add New CBlockBackColour
650                           Exit Sub
660                       End If
                          
                          Dim foreColour As Byte
                          Dim backColour As Byte
                          
670                       count = parseColourCode(text, count + 1, foreColour, backColour) - 1
                          
680                       If foreColour <> 255 Then
690                           Set blockForeColour = New CBlockForeColour
                              
700                           blockForeColour.hasForeColour = True
710                           blockForeColour.foreColour = foreColour
720                           blocks.Add blockForeColour
                              
730                       Else
740                           blocks.Add New CBlockForeColour
                              
750                           If backColour = 255 Then
760                               blocks.Add New CBlockBackColour
770                           End If
780                       End If
                          
790                       If backColour <> 255 Then
800                           Set blockbackColour = New CBlockBackColour
                              
810                           blockbackColour.hasBackColour = True
820                           blockbackColour.backColour = backColour
830                           blocks.Add blockbackColour
840                       End If
850               End Select
                  
860               last = count + 1
870           Else
880               isUrl = False
              
890               If allowArtifacts Then
900                   If wChar = AscW("h") Or wChar = AscW("H") Then
910                       If count < Len(text) Then
920                           If Mid$(LCase$(text), count, 7) = "http://" Then
930                               isUrl = True
940                           End If
950                       End If
                          
960                       If count < Len(text) Then
970                           If Mid$(LCase$(text), count, 8) = "https://" Then
980                               isUrl = True
990                           End If
1000                      End If
1010                  ElseIf wChar = AscW("w") Or wChar = AscW("W") Then
1020                      If count < Len(text) Then
1030                          If Mid$(LCase$(text), count, 4) = "www." Then
1040                              isUrl = True
1050                          End If
1060                      End If
1070                  ElseIf wChar = AscW("#") Then
1080                      isUrl = True
1090                  End If
1100              End If
                  
1110              If isUrl Then
1120                  Length = count - last
                      
1130                  If Length > 0 Then
1140                      Set blockText = New CBlockText
          
1150                      blockText.text = Mid$(text, last, Length)
1160                      blocks.Add blockText
1170                  End If
                      
1180                  count = count + parseLink(text, count, blocks)
1190                  last = count
1200              End If
1210          End If
1220      Next count
          
1230      If count - last > 0 Then
1240          If last <= Len(text) Then
1250              Set blockText = New CBlockText
              
1260              blockText.text = Mid$(text, last)
1270              blocks.Add blockText
1280          End If
1290      End If
End Sub

Private Function parseLink(ByRef text As String, index As Long, blocks As cArrayList) As Long
          Dim url As String
          Dim count As Long
          
10        For count = index To Len(text)
20            If Mid$(text, count, 1) = "'" Then
30                url = Mid$(text, index, count - index)
40                Exit For
50            ElseIf Mid$(text, count, 1) = "]" Then
60                url = Mid$(text, index, count - index)
70                Exit For
80            ElseIf Mid$(text, count, 1) = """" Then
90                url = Mid$(text, index, count - index)
100               Exit For
110           ElseIf Mid$(text, count, 1) = ")" Then
120               url = Mid$(text, index, count - index)
130               Exit For
140           ElseIf Mid$(text, count, 1) = "." Then
150               If count < Len(text) Then
160                   If Mid$(text, count + 1, 1) = " " Then
170                       url = Mid$(text, index, count - index)
180                       Exit For
190                   End If
200               Else
210                   url = Mid$(text, index, count - index)
220                   Exit For
230               End If
240           ElseIf Mid$(text, count, 1) = " " Then
250               url = Mid$(text, index, count - index)
260               Exit For
270           End If
280       Next count
          
290       If LenB(url) = 0 Then
300           url = Mid$(text, index)
310       End If
          
          Dim urlBlock As New CBlockText

320       urlBlock.isUrl = True
330       blocks.Add urlBlock
          
          Dim urlBlocks As New cArrayList
          Dim block As ITextRenderBlock
          Dim blockText As CBlockText
          
340       parseIrcFormatting url, False, urlBlocks, False
          
350       For count = 1 To urlBlocks.count
360           Set block = urlBlocks.item(count)
          
370           If TypeOf block Is CBlockText Then
380               Set blockText = block
390               urlBlock.text = urlBlock.text & blockText.text
400           Else
410               blocks.Add block
420           End If
430       Next count
          
440       urlBlock.url = urlBlock.text
          
450       parseLink = Len(url)
End Function

Private Function parseColourCode(text As String, start As Integer, ByRef fore As Byte, ByRef back _
    As Byte) As Integer
          
          Dim colourCount As Integer
          Dim digits As Byte
          Dim currentColour As Byte
          Dim hasColour As Boolean
          
          Dim wChar As Integer
          
10       On Error GoTo parseColourCode_Error

20        fore = 255
30        back = 255
          
40        For colourCount = start To Len(text)
50            wChar = AscW(Mid$(text, colourCount, 1))
              
60            If wChar > 47 And wChar < 58 Then
70                If digits = 0 Then
80                    hasColour = True
90                    currentColour = wChar - 48
100                   digits = 1
110                   start = start + 1
120               ElseIf digits = 1 Then
130                   currentColour = (currentColour * 10) + (wChar - 48)
140                   digits = 2
150                   start = start + 1
160               Else
170                   Exit For
180               End If
190           ElseIf wChar = AscW(",") Then
200               If Not hasColour Then
210                   Exit For
220               End If
                  
230               fore = currentColour
                  
240               If fore > 15 Then
250                   fore = fore Mod 16
260               End If
                  
270               hasColour = False
280               digits = 0
290               start = start + 1
300           Else
310               Exit For
320           End If
330       Next colourCount

340       If fore <> 255 Then
350           If hasColour Then
360               back = currentColour
                  
370               If back > 15 Then
380                   back = back Mod 16
390               End If
400           End If
410       ElseIf hasColour Then
420           fore = currentColour
              
430           If fore > 15 Then
440               fore = fore Mod 16
450           End If
460       End If
          
470       parseColourCode = start

480      On Error GoTo 0
490      Exit Function

parseColourCode_Error:
500       handleError "parseColourCode", Err.Number, Err.Description, Erl, vbNullString
End Function

Private Function renderTimestamp() As CBlockText
10        Set renderTimestamp = New CBlockText
          
          Dim aDate As Date
          
20        aDate = CDate(25569 + (m_timeStamp / 86400))
          
30        renderTimestamp.text = swiftFormatTime(aDate, g_timestampFormat) & " "
End Function

Private Sub addSeperator(drawingData As CDrawingData)
          Dim physLine As CPhysLine
          Dim seperator As New CBlockText
          Dim textSize As SIZE
          
10        swiftGetTextExtentPoint32 drawingData.Dc, "-", textSize
          
20        seperator.text = "-"
30        seperator.left = 0
40        seperator.width = textSize.cx
          
50        Set physLine = addPhysLine
60        physLine.addBlock seperator
End Sub

Public Function wordWrap(drawingData As CDrawingData) As Long
10        If m_wrapped Then
20            Exit Function
30        End If
          
40        m_physLines.clear
          
          Dim maxFit As Long
          Dim textSize As SIZE

          Dim break As Integer
          Dim blocks As New cArrayList

50        If timestamps Then blocks.Add renderTimestamp

60        If m_flags And TVE_NOEVENT Then
70            parseIrcFormatting m_format, True, blocks
80        Else
90            parseIrcFormatting m_eventDef.format, True, blocks
100       End If
          
          Dim blockCount As Integer
          
          Dim block As ITextRenderBlock
          
          Dim textBlock As CBlockText
          Dim newTextBlock As CBlockText
          
          Dim oldFont As Long
          
          Dim width As Long
          Dim x As Integer
          
110       width = drawingData.width
          
120       drawingData.bold = False
130       drawingData.italic = False
140       drawingData.underline = False
150       m_toggleBold = False
160       m_toggleItalic = False
170       m_toggleUnderline = False
180       m_toggleReverse = False
          
190       If seperatorTop(drawingData.ignoreSeperators) And Not seperatorAbove Then
200           addSeperator drawingData
210       End If
          
          Dim tooSmall As Boolean
          Dim noBreakTooSmall As Boolean
          Dim physLine As CPhysLine

220       Set physLine = addPhysLine
          
230       Do While blockCount < blocks.count
240           blockCount = blockCount + 1
250           Set block = blocks.item(blockCount)
              
260           If TypeOf block Is CBlockText Then
270               Set textBlock = block
              
280               swiftGetTextExtentExPoint drawingData.Dc, textBlock.text, width, maxFit, textSize
              
290               If Len(textBlock.text) > maxFit Then
300                   If maxFit <> 0 Then
310                       tooSmall = False
                      
320                       For break = maxFit To 1 Step -1
330                           If AscW(Mid$(textBlock.text, break, 1)) = 32 Then
340                               Exit For
350                           End If
360                       Next break
                          
370                       If break = 0 Then
380                           If Not noBreakTooSmall And physLine.blockCount > 0 Then
390                               blockCount = blockCount - 1
400                               noBreakTooSmall = True
410                           Else
420                               break = maxFit
430                           End If
440                       Else
450                           noBreakTooSmall = False
460                       End If
                          
470                       If break > 0 Then
480                           Set newTextBlock = New CBlockText
                              
490                           newTextBlock.text = Mid$(textBlock.text, break + 1)
500                           newTextBlock.isUrl = textBlock.isUrl
510                           newTextBlock.url = textBlock.url
                              
520                           blocks.Add newTextBlock, blockCount + 1
                              
530                           textBlock.text = Mid$(textBlock.text, 1, break)
                              
540                           textBlock.left = x
                              
550                           swiftGetTextExtentPoint32 drawingData.Dc, textBlock.text, _
                                  textSize
560                           textBlock.width = textSize.cx
                              
570                           physLine.addBlock textBlock
580                       End If
590                   Else
600                       If tooSmall Then
610                           Exit Do
620                       End If
                          
630                       tooSmall = True
640                       blockCount = blockCount - 1
650                   End If
                      
660                   Set physLine = addPhysLine
                      
670                   If indent Then
680                       width = drawingData.width - 10
690                       x = 10
700                   Else
710                       width = drawingData.width
720                       x = 0
730                   End If
                      
740                   physLine.cloneFormat drawingData
750               Else
760                   tooSmall = False
770                   noBreakTooSmall = False
780                   textBlock.left = x
790                   textBlock.width = textSize.cx
                      
800                   physLine.addBlock textBlock
                      
810                   x = x + textSize.cx
820                   width = width - textSize.cx
830               End If
840           ElseIf TypeOf block Is CBlockImage Then
850               tooSmall = False
                  
860               If width < drawingData.fontHeight Then
870                   Set physLine = addPhysLine
880                   x = drawingData.fontHeight
890                   width = drawingData.width - drawingData.fontHeight
                      
900                   block.left = 0
910                   block.width = drawingData.fontHeight
                      
920                   physLine.addBlock block
930               Else
940                   block.left = x
950                   block.width = drawingData.fontHeight
960                   width = width - drawingData.fontHeight
970                   x = x + drawingData.fontHeight
980                   physLine.addBlock block
990               End If
1000          Else
1010              tooSmall = False
              
1020              block.render drawingData
1030              physLine.addBlock block
1040          End If
1050      Loop
          
1060      If seperatorBottom(drawingData.ignoreSeperators) And Not seperatorBelow Then
1070          addSeperator drawingData
1080      End If
          
1090      wordWrap = m_physLines.count - m_oldPhysLines
          
1100      m_oldPhysLines = m_physLines.count
1110      m_wrapped = True
End Function

Public Sub clearWrap()
10        m_physLines.clear
20        m_wrapped = False
End Sub

Public Sub needsWrapping()
10        m_wrapped = False
End Sub

Public Sub render(drawingData As CDrawingData, physLine As Long)
          Dim count As Integer
          Dim start As Integer
          
10        If physLine > 0 Then
20            start = physLine
30        Else
40            start = m_physLines.count
50        End If
          
          Dim line As CPhysLine
          
          Dim blockRect As RECT
          
          
60        If m_flags And TVE_NOEVENT Then
70            If m_flags And TVE_CUSTOMIRCCOLOUR Then
80                drawingData.lineForeColour = m_foreColour
90            Else
100               drawingData.lineForeColour = m_eventColour.colour
110           End If
120       Else
130           drawingData.lineForeColour = m_eventDef.eventColour.colour
140       End If

150       m_bottom = drawingData.realY
          
160       drawingData.reset
          
170       For count = start To 1 Step -1
180           drawingData.x = 0
          
190           Set line = m_physLines.item(count)
200           line.render drawingData
210           drawingData.reset
              
220           drawingData.y = drawingData.y - drawingData.fontHeight
230           drawingData.realY = drawingData.realY - drawingData.fontHeight
              
240           If drawingData.y <= -drawingData.fontHeight Then
250               Exit For
260           End If
270       Next count

280       m_top = drawingData.realY
End Sub

Public Function getSelectedText(drawingData As CDrawingData, withFormatCodes As Boolean) As String
          Dim count As Integer
          Dim text As String
          
10        For count = 1 To m_physLines.count
20            text = text & m_physLines.item(count).getSelectedText(drawingData, withFormatCodes)
30        Next count
          
40        text = text & vbCrLf
          
50        getSelectedText = text
End Function

Public Sub shiftedUp(pixels As Long)
10        m_top = m_top - pixels
20        m_bottom = m_bottom - pixels
          
          Dim count As Integer
          
30        For count = 1 To m_physLines.count
40            m_physLines.item(count).top = m_physLines.item(count).top - pixels
50            m_physLines.item(count).bottom = m_physLines.item(count).bottom - pixels
60        Next count
End Sub

Public Sub shiftedDown(pixels As Long)
10        m_top = m_top + pixels
20        m_bottom = m_bottom + pixels
          
          Dim count As Integer
          
30        For count = 1 To m_physLines.count
40            m_physLines.item(count).top = m_physLines.item(count).top + pixels
50            m_physLines.item(count).bottom = m_physLines.item(count).bottom + pixels
60        Next count
End Sub

Private Sub Class_Initialize()
10        m_oldPhysLines = 1
20        m_timeStamp = CLng(DateDiff("s", #1/1/1970#, Now))
End Sub
