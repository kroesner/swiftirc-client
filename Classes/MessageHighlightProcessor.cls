VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageHighlightProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private emphasisPositions As cArrayList

Public Function processTextForHighlights(text As String) As String
    Dim highlightMatches As MessageHighlightMatchData

    Set highlightMatches = MessageHighlights.findHighlightMatches(m_currentNickname, text)
    
    Dim i As Long
    
    For i = 1 To highlightMatches.getHighlightMatchCount
        Debug.Print highlightMatches.getHighlightMatchPosition(i).getMatchStart & " " & highlightMatches.getHighlightMatchPosition(i).getMatchLength
    Next i
    
    If highlightMatches.getHighlightMatchCount <> 0 Then
        text = addHighlightEmphasisToText(text)
    End If
End Function

Private Function addHighlightEmphasisToText(text As String, highlightMatches As MessageHighlightMatchData)
    text = stripFormattingCodes(text)
    
    Dim i As count
    
    
    For i = 1 To highlightMatches.getHighlightMatchCount
        
    Next i
End Function

Private Sub addEmphasisPosition(start As Long, length As Long)
    Dim emphasisPosition As New HighlightMatchPositionImpl
    
    emphasisPosition.setMatchStart = start
    emphasisPosition.setMatchLength = length
    
    emphasisPositions.Add emphasisPosition
End Sub

Private Function isHightlightInsideAnother()

End Function
