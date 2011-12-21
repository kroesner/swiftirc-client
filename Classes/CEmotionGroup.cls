VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmoticonGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_emoticons As New cArrayList
Private m_prefix As String

Public Property Get prefix() As String
    prefix = m_prefix
End Property

Public Property Let prefix(newValue As String)
    m_prefix = newValue
End Property

Public Sub addEmoticon(text As String, icon As CImage)
    Dim emoticon As New CEmoticon
    
    emoticon.init text, icon
    m_emoticons.Add emoticon
End Sub

Public Function findEmoticon(ByRef text As String) As CEmoticon
    Dim count As Integer
    Dim count2 As Integer
    Dim noMatch As Boolean
    
    For count = 1 To m_emoticons.count
        For count2 = 2 To Len(m_emoticons.item(count).text)
            If Mid$(m_emoticons.item(count).text, count2, 1) <> Mid$(text, count2, 1) Then
                noMatch = True
                Exit For
            End If
        Next count2
        
        If Not noMatch Then
            Set findEmoticon = m_emoticons.item(count)
            Exit Function
        End If
        
        noMatch = False
    Next count
End Function


