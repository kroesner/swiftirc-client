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
10        prefix = m_prefix
End Property

Public Property Let prefix(newValue As String)
10        m_prefix = newValue
End Property

Public Sub addEmoticon(text As String, icon As CImage)
          Dim emoticon As New CEmoticon
          
10        emoticon.init text, icon
20        m_emoticons.Add emoticon
End Sub

Public Function findEmoticon(ByRef text As String) As CEmoticon
          Dim count As Integer
          Dim count2 As Integer
          Dim noMatch As Boolean
          
10        For count = 1 To m_emoticons.count
20            For count2 = 2 To Len(m_emoticons.item(count).text)
30                If Mid$(m_emoticons.item(count).text, count2, 1) <> Mid$(text, count2, 1) Then
40                    noMatch = True
50                    Exit For
60                End If
70            Next count2
              
80            If Not noMatch Then
90                Set findEmoticon = m_emoticons.item(count)
100               Exit Function
110           End If
              
120           noMatch = False
130       Next count
End Function


