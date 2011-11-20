VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmoticonManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_emoticonGroups As New cArrayList

Public Sub addEmoticon(text As String, icon As CImage)
          Dim emoticonGroup As CEmoticonGroup
          
10        Set emoticonGroup = findEmoticonGroup(left$(text, 1))
          
20        If Not emoticonGroup Is Nothing Then
30            emoticonGroup.addEmoticon text, icon
40        Else
50            Set emoticonGroup = New CEmoticonGroup
              
60            emoticonGroup.prefix = left$(text, 1)
70            emoticonGroup.addEmoticon text, icon
80            m_emoticonGroups.Add emoticonGroup
90        End If
End Sub

Public Function findEmoticonGroup(prefix As String) As CEmoticonGroup
          Dim count As Integer
          
10        For count = 1 To m_emoticonGroups.count
20            If m_emoticonGroups.item(count).prefix = prefix Then
30                Set findEmoticonGroup = m_emoticonGroups.item(count)
40                Exit Function
50            End If
60        Next count
End Function
