VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrigin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nickname As String
Private m_username As String
Private m_hostname As String

Public Sub parseOrigin(origin As String)
    If Not origin Like "*!*@*" Then
        m_nickname = origin
        Exit Sub
    End If
    
    Dim temp As String

    Do While LenB(origin) <> 0
        If left$(origin, 1) = "!" Then
            m_nickname = temp
            temp = vbNullString
        ElseIf left$(origin, 1) = "@" Then
            m_username = temp
            temp = vbNullString
            m_hostname = Mid$(origin, 2)
        Else
            temp = temp & left$(origin, 1)
        End If
        
        origin = Mid$(origin, 2)
    Loop
End Sub

Public Property Get nickname() As String
    nickname = m_nickname
End Property

Public Property Get username() As String
    username = m_username
End Property

Public Property Get hostname() As String
    hostname = m_hostname
End Property
