VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageHighlightMatchDataImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements MessageHighlightMatchData

Private Const TEXT_START As Long = 1
Private Const NO_MATCH As Long = 0
Private matches As cArrayList

Private lineOfText As String
Private currentNickname As String

Private Sub Class_Initialize()
    Set matches = New cArrayList
End Sub

Private Property Get MessageHighlightMatchData_isMatchingHighlight() As Boolean
    MessageHighlightMatchData_isMatchingHighlight = matches.count <> 0
End Property

Private Property Get MessageHighlightMatchData_getHighlightMatchCount() As Long
    MessageHighlightMatchData_getHighlightMatchCount = matches.count
End Property

Private Property Get MessageHighlightMatchData_getHighlightMatchPosition(i As Long) As HighlightMatchPosition
    Set MessageHighlightMatchData_getHighlightMatchPosition = matches.item(i)
End Property

Public Sub storeHighlightMatches(nickname As String, text As String)
    currentNickname = nickname
    lineOfText = text
    
    If settings.enableHighlighting Then
        storeHighlightNickMatch
        storeHighlightTextMatches
    End If
End Sub

Private Sub storeHighlightNickMatch()
    If settings.highlightNickname Then
        Dim position As Long
    
        position = findHighlightPositionInText(currentNickname)
        addHighlightMatch position, currentNickname
    End If
End Sub

Private Sub storeHighlightTextMatches()
    Dim i As Long
    
    For i = 1 To highlights.highlightCount
        Dim position As Long
        Dim highlightText As String
        
        highlightText = highlights.highlightItem(i).text
        position = findHighlightPositionInText(highlightText)
        addHighlightMatch position, highlightText
    Next i
End Sub

Private Function findHighlightPositionInText(highlight As String) As Long
    findHighlightPositionInText = InStr(TEXT_START, lineOfText, highlight, vbTextCompare)
End Function

Private Sub addHighlightMatch(position As Long, highlight As String)
    If position <> NO_MATCH Then
        Dim newMatchPosition As HighlightMatchPosition
        
        Set newMatchPosition = createNewMatchPosition(position, highlight)
        matches.Add newMatchPosition
    End If
End Sub

Private Function createNewMatchPosition(position, highlight As String) As HighlightMatchPosition
    Dim newMatchPosition As New HighlightMatchPositionImpl
    
    newMatchPosition.setMatchStart = position
    newMatchPosition.setMatchLength = Len(highlight)
    
    Set createNewMatchPosition = newMatchPosition
End Function

