Attribute VB_Name = "MessageHighlights"
Option Explicit

Public Type MatchPosition
    matchStart As Long
    matchLength As Long
End Type

Public Function findHighlightMatches(currentNickname As String, text As String) As MessageHighlightMatchData
    Dim matchData As New MessageHighlightMatchDataImpl
    
    matchData.storeHighlightMatches currentNickname, text
    Set findHighlightMatches = matchData
End Function
