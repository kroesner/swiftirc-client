VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HighlightMatchPositionImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements HighlightMatchPosition

Private matchStart As Long
Private matchLength As Long

Private Property Get HighlightMatchPosition_getMatchStart() As Long
    HighlightMatchPosition_getMatchStart = matchStart
End Property

Private Property Get HighlightMatchPosition_getMatchLength() As Long
    HighlightMatchPosition_getMatchLength = matchLength
End Property

Public Property Let setMatchStart(newMatchStart As Long)
    matchStart = newMatchStart
End Property

Public Property Let setMatchLength(newMatchLength As Long)
    matchLength = newMatchLength
End Property

