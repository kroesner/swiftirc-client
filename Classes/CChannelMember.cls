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
    nickname = m_nickname
End Property

Public Property Let nickname(newValue As String)
    m_nickname = newValue
End Property

Public Property Get modes() As String
    Dim count As Long
    
    For count = 1 To m_modes.count
        modes = modes & m_modes.item(count).mode
    Next count
End Property

Public Property Let modes(newValue As String)
    Dim count As Long
    Dim mode As CMode
    
    For count = 1 To Len(newValue)
        Set mode = New CMode
        mode.mode = Mid$(newValue, count, 1)
        m_modes.Add mode
    Next count
End Property

Public Sub addMode(mode As String)
    Dim newMode As New CMode
    
    newMode.mode = mode
    m_modes.Add newMode
End Sub

Public Sub removeMode(mode As String)
    Dim count As Long
    
    For count = 1 To m_modes.count
        If m_modes.item(count).mode = mode Then
            m_modes.Remove count
            Exit Sub
        End If
    Next count
End Sub

Public Function hasMode(mode As String) As Boolean
    If m_modes.count = 0 Then Exit Function
    
    Dim count As Long
    
    For count = 1 To m_modes.count
        If m_modes.item(count).mode = mode Then
            hasMode = True
            Exit Function
        End If
    Next count
End Function

