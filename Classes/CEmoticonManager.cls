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
    
    Set emoticonGroup = findEmoticonGroup(left$(text, 1))
    
    If Not emoticonGroup Is Nothing Then
        emoticonGroup.addEmoticon text, icon
    Else
        Set emoticonGroup = New CEmoticonGroup
        
        emoticonGroup.prefix = left$(text, 1)
        emoticonGroup.addEmoticon text, icon
        m_emoticonGroups.Add emoticonGroup
    End If
End Sub

Public Function findEmoticonGroup(prefix As String) As CEmoticonGroup
    Dim count As Integer
    
    For count = 1 To m_emoticonGroups.count
        If m_emoticonGroups.item(count).prefix = prefix Then
            Set findEmoticonGroup = m_emoticonGroups.item(count)
            Exit Function
        End If
    Next count
End Function
