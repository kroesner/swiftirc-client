VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChannelMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nickname As String
Private m_modes As New cArrayList

Public Property Get nickname() As String
10        nickname = m_nickname
End Property

Public Property Let nickname(newValue As String)
10        m_nickname = newValue
End Property

Public Property Get modes() As String
          Dim count As Long
          
10        For count = 1 To m_modes.count
20            modes = modes & m_modes.item(count).mode
30        Next count
End Property

Public Property Let modes(newValue As String)
          Dim count As Long
          Dim mode As CMode
          
10        For count = 1 To Len(newValue)
20            Set mode = New CMode
30            mode.mode = Mid$(newValue, count, 1)
40            m_modes.Add mode
50        Next count
End Property

Public Sub addMode(mode As String)
          Dim newMode As New CMode
          
10        newMode.mode = mode
20        m_modes.Add newMode
End Sub

Public Sub removeMode(mode As String)
          Dim count As Long
          
10        For count = 1 To m_modes.count
20            If m_modes.item(count).mode = mode Then
30                m_modes.Remove count
40                Exit Sub
50            End If
60        Next count
End Sub

Public Function hasMode(mode As String) As Boolean
10        If m_modes.count = 0 Then Exit Function
          
          Dim count As Long
          
20        For count = 1 To m_modes.count
30            If m_modes.item(count).mode = mode Then
40                hasMode = True
50                Exit Function
60            End If
70        Next count
End Function

